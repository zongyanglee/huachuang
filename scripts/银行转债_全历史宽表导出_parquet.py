from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List

import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from 底稿更新 import load_original_data


@dataclass(frozen=True)
class RunConfig:
    parquet_root: str = "data/转债个券历史序列"
    output_xlsx: str = "银行转债_全历史_分sheet宽表_parquet.xlsx"

    total_sheet: str = "总表"
    industry_col: str = "申万行业"
    industry_value: str = "银行"

    industry_map_xlsx: str | None = None
    industry_map_code_col: str = "转债代码"
    industry_map_industry_col: str = "中信行业"


def _date_cols(df: pd.DataFrame) -> List[pd.Timestamp]:
    return sorted([c for c in df.columns if isinstance(c, pd.Timestamp)])


def _safe_sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(name)).strip()
    if not base:
        base = "sheet"
    base = base[:31]
    if base not in used:
        used.add(base)
        return base
    i = 2
    while True:
        suffix = f"_{i}"
        cand = (base[: 31 - len(suffix)] + suffix)[:31]
        if cand not in used:
            used.add(cand)
            return cand
        i += 1


def _iter_metric_sheets(data: dict[str, pd.DataFrame]) -> Iterable[str]:
    for k, df in data.items():
        if not isinstance(df, pd.DataFrame):
            continue
        if not _date_cols(df):
            continue
        yield k


def _select_codes(config: RunConfig, data: dict[str, pd.DataFrame]) -> tuple[list[str], str]:
    if config.industry_map_xlsx:
        map_path = Path(config.industry_map_xlsx).expanduser().resolve()
        if not map_path.exists():
            raise FileNotFoundError(f"未找到行业映射文件: {map_path}")
        mdf = pd.read_excel(map_path)
        for col in (config.industry_map_code_col, config.industry_map_industry_col):
            if col not in mdf.columns:
                raise KeyError(f"行业映射缺少列: {col}，实际列={list(mdf.columns)}")
        codes = (
            mdf.loc[
                mdf[config.industry_map_industry_col].astype("object").eq(config.industry_value),
                config.industry_map_code_col,
            ]
            .astype("string")
            .dropna()
            .drop_duplicates()
            .tolist()
        )
        if not codes:
            raise ValueError(f"未筛选到转债：{map_path.name}.{config.industry_map_industry_col} == {config.industry_value}")
        return codes, f"{map_path.name}.{config.industry_map_industry_col} == {config.industry_value}"

    if config.total_sheet not in data:
        raise KeyError(f"parquet 数据缺少 sheet: {config.total_sheet}")
    total_df = data[config.total_sheet].copy()
    if "__row_id" in total_df.columns:
        total_df = total_df.set_index("__row_id")
    industry_s = total_df.get(config.industry_col)
    if industry_s is None:
        raise KeyError(f"`{config.total_sheet}` 中未找到 `{config.industry_col}` 字段。")
    mask = industry_s.astype("object").eq(config.industry_value)
    codes = mask[mask].index.tolist()
    if not codes:
        raise ValueError(f"未筛选到转债：{config.total_sheet}.{config.industry_col} == {config.industry_value}")
    return codes, f"{config.total_sheet}.{config.industry_col} == {config.industry_value}"


def main(config: RunConfig) -> Path:
    data = load_original_data(source_type="parquet", parquet_root=config.parquet_root, force_refresh=False)
    codes, filter_note = _select_codes(config, data)

    out_path = Path(config.output_xlsx).resolve()
    used_sheet_names: set[str] = set()
    sheet_map_rows: list[dict[str, str]] = []

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        info = pd.DataFrame(
            [
                {"key": "数据源", "value": f"parquet: {Path(config.parquet_root).resolve()}"},
                {"key": "筛选口径", "value": filter_note},
                {"key": "转债数量(静态)", "value": int(len(codes))},
                {"key": "列含义", "value": "每列=转债代码，每行=日期；值来自对应指标sheet"},
            ]
        )
        info.to_excel(writer, sheet_name=_safe_sheet_name("说明", used_sheet_names), index=False)

        for metric in _iter_metric_sheets(data):
            df = data[metric]
            if "__row_id" in df.columns:
                df = df.set_index("__row_id")

            dates = _date_cols(df)
            if not dates:
                continue

            sub = df.reindex(codes)
            wide = sub[dates].T
            wide.index.name = "日期"

            sheet_name = _safe_sheet_name(metric, used_sheet_names)
            wide.to_excel(writer, sheet_name=sheet_name, index=True)
            sheet_map_rows.append({"metric": metric, "sheet_name": sheet_name})

        if sheet_map_rows:
            pd.DataFrame(sheet_map_rows).to_excel(
                writer, sheet_name=_safe_sheet_name("sheet映射", used_sheet_names), index=False
            )

    return out_path


def _parse_args() -> RunConfig:
    p = argparse.ArgumentParser(description="从parquet导出指定行业转债的全历史宽表（按指标分sheet）")
    p.add_argument("--parquet-root", default=RunConfig.parquet_root)
    p.add_argument("--industry", default=RunConfig.industry_value, help="筛选的行业值")
    p.add_argument("--output", default=RunConfig.output_xlsx, help="输出xlsx路径")
    p.add_argument("--industry-map", default=None, help="行业映射xlsx（优先使用，不依赖parquet总表行业字段）")
    p.add_argument("--industry-map-code-col", default=RunConfig.industry_map_code_col)
    p.add_argument("--industry-map-industry-col", default=RunConfig.industry_map_industry_col)
    args = p.parse_args()
    return RunConfig(
        parquet_root=args.parquet_root,
        industry_value=args.industry,
        output_xlsx=args.output,
        industry_map_xlsx=args.industry_map,
        industry_map_code_col=args.industry_map_code_col,
        industry_map_industry_col=args.industry_map_industry_col,
    )


if __name__ == "__main__":
    out = main(_parse_args())
    print(f"[ok] wrote: {out}")

