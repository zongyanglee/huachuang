from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class RunConfig:
    parquet_root: Path = Path("data/转债个券历史序列")
    output_dir: Path = Path("0531数据更新")
    output_stem: str = "价格分类溢价率_le110_120_130_140"

    parity_sheet: str = "平价"
    conv_premium_sheet: str = "转股溢价率"
    purebond_premium_sheet: str = "纯债溢价率"

    # 分组口径（用户自定义）：<=110, (110,120], (120,130], (130,140], >140
    bins: Tuple[Tuple[float, float], ...] = (
        (-np.inf, 110),
        (110, 120),
        (120, 130),
        (130, 140),
        (140, np.inf),
    )
    labels: Tuple[str, ...] = ("<=110", "110-120", "120-130", "130-140", "140+")


def _find_latest_monthly_parquet(parquet_root: Path) -> Path:
    if not parquet_root.exists():
        raise FileNotFoundError(f"parquet_root 不存在: {parquet_root}")

    monthly_files: List[Path] = []
    for p in parquet_root.rglob("*.parquet"):
        # 仅选择 年份目录/yyyymm.parquet，排除 _special/_meta 等
        if not p.parent.name.isdigit():
            continue
        if not re.match(r"^\d{6}\.parquet$", p.name):
            continue
        monthly_files.append(p)

    if not monthly_files:
        raise FileNotFoundError(f"未在 {parquet_root} 下找到 yyyy/yyyymm.parquet 结构的月度文件")

    return max(monthly_files, key=lambda x: x.stat().st_mtime)


def _sheet_wide(monthly_df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    sub = monthly_df.loc[monthly_df["__sheet_name"] == sheet_name].copy()
    if sub.empty:
        raise KeyError(f"月度parquet中缺少sheet: {sheet_name}")

    sub = sub.drop(columns=["__sheet_name"]).set_index("__row_id")
    sub.index.name = None

    col_map: Dict[object, pd.Timestamp] = {}
    for c in sub.columns:
        ts = pd.to_datetime(c, errors="coerce")
        if pd.notna(ts):
            col_map[c] = pd.Timestamp(ts)

    date_cols = [c for c in sub.columns if c in col_map]
    if not date_cols:
        raise ValueError(f"sheet={sheet_name} 未找到可解析为日期的列")

    wide = sub[date_cols].copy()
    wide.columns = [col_map[c] for c in date_cols]
    wide = wide.sort_index(axis=1)
    return wide


def _as_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def _bucket_mean(
    parity_s: pd.Series,
    metric_s: pd.Series,
    bins: Sequence[Tuple[float, float]],
    labels: Sequence[str],
) -> Dict[str, float]:
    x = _as_num(parity_s)
    y = _as_num(metric_s)
    out: Dict[str, float] = {}
    for (lo, hi), lb in zip(bins, labels):
        if np.isneginf(lo):
            mask = x.le(hi)
        elif np.isposinf(hi):
            mask = x.gt(lo)
        else:
            # 与历史notebook一致：左开右闭 (80, 90] 这类区间
            mask = x.gt(lo) & x.le(hi)
        mask = mask & y.notna()
        out[lb] = float(y[mask].mean()) if mask.any() else float("nan")
    return out


def _calc_bucket_table(
    parity_df: pd.DataFrame,
    metric_df: pd.DataFrame,
    metric_name: str,
    bins: Sequence[Tuple[float, float]],
    labels: Sequence[str],
) -> pd.DataFrame:
    common_dates = sorted(set(parity_df.columns) & set(metric_df.columns))
    if not common_dates:
        raise ValueError(f"{metric_name}: 平价与指标没有共同日期列")

    rows: List[dict] = []
    for d in common_dates:
        row = {"日期": d}
        row.update(_bucket_mean(parity_df[d], metric_df[d], bins=bins, labels=labels))
        rows.append(row)

    out = pd.DataFrame(rows).set_index("日期").sort_index()
    out = out.reindex(columns=list(labels))
    return out


def main(cfg: RunConfig) -> Path:
    latest_parquet = _find_latest_monthly_parquet(cfg.parquet_root)
    monthly_df = pd.read_parquet(latest_parquet)

    parity_df = _sheet_wide(monthly_df, cfg.parity_sheet)
    conv_premium_df = _sheet_wide(monthly_df, cfg.conv_premium_sheet)
    purebond_premium_df = _sheet_wide(monthly_df, cfg.purebond_premium_sheet)

    conv_table = _calc_bucket_table(parity_df, conv_premium_df, "价格分类转股溢价率", bins=cfg.bins, labels=cfg.labels)
    pure_table = _calc_bucket_table(parity_df, purebond_premium_df, "价格分类纯债溢价率", bins=cfg.bins, labels=cfg.labels)

    cfg.output_dir.mkdir(parents=True, exist_ok=True)
    out_xlsx = cfg.output_dir / f"{cfg.output_stem}_{latest_parquet.stem}.xlsx"
    with pd.ExcelWriter(out_xlsx, engine="openpyxl", mode="w") as writer:
        conv_table.to_excel(writer, sheet_name="价格分类转股溢价率")
        pure_table.to_excel(writer, sheet_name="价格分类纯债溢价率")

    return out_xlsx


if __name__ == "__main__":
    out = main(RunConfig())
    print(f"[ok] wrote: {out}")
