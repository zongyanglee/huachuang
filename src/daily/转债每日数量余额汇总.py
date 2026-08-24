# -*- coding: utf-8 -*-
"""
读取目录下（递归）的 Parquet 文件，汇总“余额”指标：
- 每日转债个数：当日余额 > 0 的转债数量
- 每日总余额：当日余额求和

默认读取：./转债个券历史序列/**/*.parquet（递归，跳过 _meta）
输出文件：./转债每日数量与总余额汇总.xlsx（sheet=daily）
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

import sys

_COMMON_MODULE_DIR = Path(__file__).resolve().parents[1] / "common"
if str(_COMMON_MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(_COMMON_MODULE_DIR))

from 转债Parquet标准读写模块 import BOND_CODE, TRADE_DATE


def iter_parquet_files(root: Path) -> list[Path]:
    files = [p for year in root.iterdir() if year.is_dir() and year.name.isdigit() for p in year.glob("*.parquet")]
    files.sort(key=lambda p: str(p))
    return files


def _normalize_date_columns(columns: list[object]) -> dict[object, pd.Timestamp]:
    mapping: dict[object, pd.Timestamp] = {}
    for c in columns:
        if c in ("__sheet_name", "__row_id"):
            continue
        try:
            ts = pd.to_datetime(c)
        except Exception:
            continue
        if pd.isna(ts):
            continue
        mapping[c] = ts
    return mapping


def aggregate_one_file(path: Path, sheet_name: str = "余额") -> pd.DataFrame:
    df = pd.read_parquet(path, columns=[BOND_CODE, TRADE_DATE, sheet_name])
    if df.empty or sheet_name not in df:
        return pd.DataFrame(columns=["count", "total_balance", "source_file"])
    df[TRADE_DATE] = pd.to_datetime(df[TRADE_DATE])
    values = pd.to_numeric(df[sheet_name], errors="coerce")
    out = pd.DataFrame({"date": df[TRADE_DATE], "value": values})
    out = out.groupby("date", as_index=False).agg(
        count=("value", lambda series: int((series > 0).sum())),
        total_balance=("value", "sum"),
    )
    out["source_file"] = str(path)
    out = out.sort_values("date")
    out = out.set_index("date")
    out.index.name = "date"
    return out


def aggregate_all(parquet_root: Path, sheet_name: str = "余额") -> pd.DataFrame:
    files = iter_parquet_files(parquet_root)
    if not files:
        raise FileNotFoundError(f"未找到 parquet：{parquet_root}")

    parts = [aggregate_one_file(p, sheet_name=sheet_name) for p in files]
    all_days = pd.concat(parts, axis=0)
    all_days.index.name = "date"

    merged = (
        all_days.reset_index()
        .sort_values(["date", "source_file"])
        .groupby("date", as_index=False)
        .agg({"count": "max", "total_balance": "max", "source_file": "last"})
        .set_index("date")
        .sort_index()
    )
    merged.index.name = "date"
    return merged


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--parquet-root",
        default=str(Path("data/转债个券历史序列")),
        help="Parquet 根目录（递归搜索 *.parquet）",
    )
    parser.add_argument("--sheet-name", default="余额", help="指标名（__sheet_name）")
    parser.add_argument(
        "--out",
        default=str(Path("转债每日数量与总余额汇总.xlsx")),
        help="输出 Excel 路径",
    )
    args = parser.parse_args()

    daily = aggregate_all(Path(args.parquet_root), sheet_name=args.sheet_name)

    daily_out = daily.copy()
    daily_out.insert(0, "date", daily_out.index)
    daily_out["date"] = pd.to_datetime(daily_out["date"]).dt.date

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        daily_out.to_excel(writer, sheet_name="daily", index=False)

    print(f"已输出：{out_path.resolve()}")
    print(f"行数：{len(daily_out)}  日期范围：{daily.index.min()} ~ {daily.index.max()}")


if __name__ == "__main__":
    main()
