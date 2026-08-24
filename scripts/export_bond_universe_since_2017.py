#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export the unique convertible-bond universe overlapping 2017 onward."""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "outputs" / "bond_universe_since_2017"
CUTOFF = pd.Timestamp("2017-01-01")


def parse_date(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce")
    return parsed.where(parsed.dt.year >= 1990)


def observed_active_codes_by_year(start_year: int, end_year: int) -> dict[int, set[str]]:
    result: dict[int, set[str]] = {}
    for year in range(start_year, end_year + 1):
        active: set[str] = set()
        for file in sorted((DATA_ROOT / str(year)).glob("*.parquet")):
            frame = pd.read_parquet(file, filters=[("__sheet_name", "==", "余额")])
            date_cols = [c for c in frame.columns if c not in {"__sheet_name", "__row_id"}]
            numeric = frame[date_cols].apply(pd.to_numeric, errors="coerce")
            active.update(frame.loc[numeric.gt(0).any(axis=1), "__row_id"].astype(str))
        result[year] = active
    return result


def latest_snapshot() -> tuple[pd.Timestamp, dict[str, pd.Series]]:
    files = sorted((DATA_ROOT / "2026").glob("*.parquet"))
    if not files:
        raise FileNotFoundError("未找到2026年历史分片")
    frame = pd.read_parquet(files[-1])
    date_cols = [c for c in frame.columns if c not in {"__sheet_name", "__row_id"}]
    parsed = pd.to_datetime(pd.Index(date_cols), errors="coerce")
    valid_pairs = [(col, pd.Timestamp(ts)) for col, ts in zip(date_cols, parsed) if pd.notna(ts)]
    latest_col, latest_date = max(valid_pairs, key=lambda item: item[1])
    metrics: dict[str, pd.Series] = {}
    for metric in ["余额", "收盘价", "债项评级", "主体评级", "交易状态"]:
        sub = frame.loc[frame["__sheet_name"].eq(metric), ["__row_id", latest_col]].copy()
        metrics[metric] = sub.set_index("__row_id")[latest_col]
    return latest_date, metrics


def year_end_balance_count(year: int) -> tuple[pd.Timestamp, int]:
    files = sorted((DATA_ROOT / str(year)).glob("*.parquet"))
    if not files:
        return pd.NaT, 0
    frame = pd.read_parquet(files[-1])
    date_cols = [c for c in frame.columns if c not in {"__sheet_name", "__row_id"}]
    parsed = pd.to_datetime(pd.Index(date_cols), errors="coerce")
    valid_pairs = [(col, pd.Timestamp(ts)) for col, ts in zip(date_cols, parsed) if pd.notna(ts)]
    latest_col, snapshot_date = max(valid_pairs, key=lambda item: item[1])
    balance = frame.loc[frame["__sheet_name"].eq("余额"), latest_col]
    return snapshot_date, int(pd.to_numeric(balance, errors="coerce").gt(0).sum())


def clean_text(value: object) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    return None if text in {"", "0", "nan", "None", "<NA>"} else text


def main() -> None:
    total = pd.read_parquet(DATA_ROOT / "_special" / "总表.parquet").copy()
    total["上市日期_解析"] = parse_date(total["上市日期"])
    total["发行日期_解析"] = parse_date(total["发行日期"])
    total["最后交易日_解析"] = parse_date(total["最后交易日"])
    total["存续起始日"] = total["上市日期_解析"].fillna(total["发行日期_解析"])

    latest_date, latest = latest_snapshot()
    active_by_year = observed_active_codes_by_year(CUTOFF.year, latest_date.year)
    observed_codes = set().union(*active_by_year.values())
    include = total["__row_id"].astype(str).isin(observed_codes)
    universe = total.loc[include].copy()
    universe["最新数据日"] = latest_date
    for metric, series in latest.items():
        universe[f"最新{metric}"] = universe["__row_id"].map(series)

    universe["最新余额"] = pd.to_numeric(universe["最新余额"], errors="coerce")
    universe["最新收盘价"] = pd.to_numeric(universe["最新收盘价"], errors="coerce")
    universe["截至最新数据日有余额"] = universe["最新余额"].gt(0)
    universe["截至最新数据日可交易"] = (
        universe["截至最新数据日有余额"]
        & universe["最新交易状态"].astype("string").isin(["交易", "新股上市"])
        & universe["最新收盘价"].notna()
    )
    universe["起始日期说明"] = np.where(
        universe["上市日期_解析"].notna(), "上市日期", "上市日期缺失，采用发行日期"
    )
    universe["备注"] = ""
    universe.loc[
        universe["最后交易日_解析"].gt(latest_date), "备注"
    ] = "最后交易日晚于最新数据日，是否当前存续以最新余额判断"
    universe.loc[
        universe["截至最新数据日有余额"] & ~universe["截至最新数据日可交易"], "备注"
    ] = "最新仍有余额，但非正常交易状态"

    output_columns = [
        "__row_id", "转债名称", "上市日期_解析", "最后交易日_解析", "发行日期_解析",
        "发行规模", "申万行业", "转股期起始日", "赎回公告日", "赎回触发比例",
        "赎回触发计算时间区间", "赎回触发计算最大时间区间", "下修触发比例",
        "重设触发计算时间区间", "重设触发计算最大时间区间", "最新数据日",
        "最新余额", "最新收盘价", "最新债项评级", "最新主体评级", "最新交易状态",
        "截至最新数据日有余额", "截至最新数据日可交易", "起始日期说明", "备注",
    ]
    universe = universe[output_columns].rename(
        columns={
            "__row_id": "转债代码",
            "上市日期_解析": "上市日期",
            "最后交易日_解析": "最后交易日",
            "发行日期_解析": "发行日期",
        }
    )
    universe = universe.sort_values(["上市日期", "转债代码"], ascending=[True, True]).reset_index(drop=True)

    annual_rows = []
    start_dates = universe["上市日期"].fillna(universe["发行日期"])
    end_dates = universe["最后交易日"]
    for year in range(2017, latest_date.year + 1):
        year_start = pd.Timestamp(year=year, month=1, day=1)
        year_end = min(pd.Timestamp(year=year, month=12, day=31), latest_date)
        overlap = universe["转债代码"].isin(active_by_year[year])
        new_listings = start_dates.between(year_start, year_end)
        exits = end_dates.between(year_start, year_end)
        year_end_active = start_dates.le(year_end) & (end_dates.isna() | end_dates.ge(year_end))
        snapshot_date, balance_count = year_end_balance_count(year)
        annual_rows.append(
            {
                "年份": year,
                "年内曾有余额数量": int(overlap.sum()),
                "当年上市数量": int(new_listings.sum()),
                "当年结束交易数量": int(exits.sum()),
                "年末日期区间存续数量": int(year_end_active.sum()),
                "年末快照日期": snapshot_date,
                "年末有余额数量": balance_count,
            }
        )
    annual = pd.DataFrame(annual_rows)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    universe.to_csv(OUTPUT_DIR / "2017年以来全部存续转债列表.csv", index=False, encoding="utf-8-sig")
    payload = {
        "cutoff": str(CUTOFF.date()),
        "latest_date": str(latest_date.date()),
        "total_count": int(len(universe)),
        "current_balance_count": int(universe["截至最新数据日有余额"].sum()),
        "current_tradable_count": int(universe["截至最新数据日可交易"].sum()),
        "missing_listing_count": int((universe["起始日期说明"] != "上市日期").sum()),
        "columns": list(universe.columns),
        "rows": universe.replace({np.nan: None, pd.NaT: None}).to_dict(orient="records"),
        "annual": annual.to_dict(orient="records"),
    }
    (OUTPUT_DIR / "universe.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    print(f"最新数据日: {latest_date:%Y-%m-%d}")
    print(f"2017年以来曾存续转债: {len(universe)}只")
    print(f"最新数据日有余额: {universe['截至最新数据日有余额'].sum()}只")
    print(annual.to_string(index=False))


if __name__ == "__main__":
    main()
