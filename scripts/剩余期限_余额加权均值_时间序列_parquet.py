from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class RunConfig:
    parquet_root: str = "data/转债个券历史序列"
    output_dir: str = "outputs/remaining_maturity_balance_weighted"
    output_stem: str = "历史余额加权剩余期限均值序列"

    balance_sheet: str = "余额"
    remaining_maturity_sheet: str = "剩余期限"


def _monthly_parquet_files(parquet_root: Path) -> List[Path]:
    files = [
        p
        for p in parquet_root.glob("*/*.parquet")
        if p.parent.name.isdigit() and p.stem.isdigit()
    ]
    return sorted(files)


def _date_columns(columns: Iterable[object]) -> Dict[object, pd.Timestamp]:
    mapping: Dict[object, pd.Timestamp] = {}
    for col in columns:
        if col in {"__sheet_name", "__row_id"}:
            continue
        ts = pd.to_datetime(col, errors="coerce")
        if pd.notna(ts):
            mapping[col] = pd.Timestamp(ts).normalize()
    return mapping


def _load_sheet_wide(parquet_root: Path, sheet_name: str) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    for file_path in _monthly_parquet_files(parquet_root):
        df = pd.read_parquet(file_path)
        part = df[df["__sheet_name"].eq(sheet_name)].copy()
        if part.empty:
            continue

        date_col_map = _date_columns(part.columns)
        if not date_col_map:
            continue

        part = part[["__row_id", *date_col_map.keys()]].rename(columns=date_col_map)
        part = part.set_index("__row_id")
        frames.append(part)

    if not frames:
        raise KeyError(f"未在 parquet 目录中找到 sheet: {sheet_name}")

    wide = pd.concat(frames, axis=1)
    wide = wide.loc[:, ~wide.columns.duplicated(keep="last")]
    wide = wide.reindex(sorted(wide.columns), axis=1)
    return wide


def _as_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _weighted_mean(value: pd.Series, weight: pd.Series) -> float:
    v = _as_num(value)
    w = _as_num(weight)
    ok = v.notna() & w.notna() & np.isfinite(v) & np.isfinite(w) & (w > 0)
    if not ok.any():
        return float("nan")

    weight_sum = float(w[ok].sum())
    if weight_sum == 0.0:
        return float("nan")

    return float((v[ok] * w[ok]).sum() / weight_sum)


def _valid_count(value: pd.Series, weight: pd.Series) -> int:
    v = _as_num(value)
    w = _as_num(weight)
    ok = v.notna() & w.notna() & np.isfinite(v) & np.isfinite(w) & (w > 0)
    return int(ok.sum())


def _weight_sum(value: pd.Series, weight: pd.Series) -> float:
    v = _as_num(value)
    w = _as_num(weight)
    ok = v.notna() & w.notna() & np.isfinite(v) & np.isfinite(w) & (w > 0)
    if not ok.any():
        return float("nan")
    return float(w[ok].sum())


def build_series(config: RunConfig) -> pd.DataFrame:
    parquet_root = Path(config.parquet_root)
    if not parquet_root.exists():
        raise FileNotFoundError(f"未找到 parquet 目录: {parquet_root.resolve()}")

    balance_df = _load_sheet_wide(parquet_root, config.balance_sheet)
    maturity_df = _load_sheet_wide(parquet_root, config.remaining_maturity_sheet)

    dates = sorted(set(balance_df.columns).intersection(maturity_df.columns))
    if not dates:
        raise ValueError("余额与剩余期限没有可共同计算的日期列")

    rows = []
    for date_col in dates:
        balance = balance_df[date_col]
        maturity = maturity_df[date_col]
        rows.append(
            {
                "日期": date_col,
                "有效个券数": _valid_count(maturity, balance),
                "有效余额合计": _weight_sum(maturity, balance),
                "剩余期限_余额加权均值": _weighted_mean(maturity, balance),
            }
        )

    return pd.DataFrame(rows).set_index("日期").sort_index()


def main(config: RunConfig = RunConfig()) -> Path:
    out_df = build_series(config)

    out_dir = Path(config.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_xlsx = (out_dir / f"{config.output_stem}.xlsx").resolve()
    out_csv = out_xlsx.with_suffix(".csv")

    info = pd.DataFrame(
        [
            {"key": "数据源", "value": f"parquet: {Path(config.parquet_root).resolve()}"},
            {"key": "取值字段", "value": config.remaining_maturity_sheet},
            {"key": "权重字段", "value": f"{config.balance_sheet}，同日余额>0"},
            {"key": "计算公式", "value": "sum(剩余期限 * 余额) / sum(余额)"},
            {
                "key": "日期范围",
                "value": f"{out_df.index.min().date()} ~ {out_df.index.max().date()} (n={len(out_df)})",
            },
        ]
    )

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        info.to_excel(writer, sheet_name="说明", index=False)
        out_df.to_excel(writer, sheet_name="时间序列", index=True)

    out_df.to_csv(out_csv, encoding="utf-8-sig")
    return out_xlsx


if __name__ == "__main__":
    output_path = main()
    print(f"[ok] wrote: {output_path}")
