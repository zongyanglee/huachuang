#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Annual convertible-bond rating structure and simple-average returns since 2018."""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "outputs" / "rating_annual_stats_since_2018"
START_YEAR = 2018
RATING_ORDER = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "未评级"]


def clean_rating(frame: pd.DataFrame) -> pd.DataFrame:
    return (
        frame.astype("string")
        .replace({"0": pd.NA, "": pd.NA, "nan": pd.NA, "None": pd.NA})
        .apply(lambda col: col.str.strip().str.upper().str.replace("＋", "+").str.replace("－", "-"))
    )


def read_year_metric(year: int, metric: str) -> pd.DataFrame:
    parts: list[pd.DataFrame] = []
    for file in sorted((DATA_ROOT / str(year)).glob("*.parquet")):
        frame = pd.read_parquet(file, filters=[("__sheet_name", "==", metric)])
        date_cols = [c for c in frame.columns if c not in {"__sheet_name", "__row_id"}]
        parsed = pd.to_datetime(pd.Index(date_cols), errors="coerce")
        keep = [col for col, date in zip(date_cols, parsed) if pd.notna(date)]
        if not keep:
            continue
        wide = frame.set_index("__row_id")[keep].T
        wide.index = pd.to_datetime(wide.index)
        wide.columns = wide.columns.astype(str)
        parts.append(wide)
    if not parts:
        return pd.DataFrame()
    result = pd.concat(parts).sort_index()
    return result.loc[~result.index.duplicated(keep="last")]


def first_valid_rating(rating: pd.Series, valid: pd.Series) -> str:
    values = rating.where(valid).dropna()
    return str(values.iloc[0]) if not values.empty else "未评级"


def make_group_stats(detail: pd.DataFrame, period_label: str) -> pd.DataFrame:
    count_total = len(detail)
    size_total = detail["发行规模"].sum(min_count=1)
    rows = []
    ratings = [rating for rating in RATING_ORDER if rating in set(detail["评级"])]
    extras = sorted(set(detail["评级"]) - set(RATING_ORDER))
    for rating in [*ratings, *extras]:
        group = detail.loc[detail["评级"].eq(rating)]
        valid_returns = group["收益率"].dropna()
        rows.append(
            {
                "期间": period_label,
                "评级": rating,
                "发行规模合计(亿元)": float(group["发行规模"].sum(min_count=1)),
                "发行规模占比": float(group["发行规模"].sum(min_count=1) / size_total) if size_total else np.nan,
                "个数": int(len(group)),
                "个数占比": float(len(group) / count_total) if count_total else np.nan,
                "收益率简单平均": float(valid_returns.mean()) if not valid_returns.empty else np.nan,
                "收益率中位数": float(valid_returns.median()) if not valid_returns.empty else np.nan,
                "收益率有效样本数": int(len(valid_returns)),
            }
        )
    return pd.DataFrame(rows)


def main() -> None:
    total = pd.read_parquet(DATA_ROOT / "_special" / "总表.parquet").set_index("__row_id")
    issue_size = pd.to_numeric(total["发行规模"], errors="coerce")
    name = total["转债名称"].astype("string")
    industry = total["申万行业"].astype("string")

    year_details: list[pd.DataFrame] = []
    current_year = max(int(p.name) for p in DATA_ROOT.iterdir() if p.is_dir() and p.name.isdigit())
    for year in range(START_YEAR, current_year + 1):
        close = read_year_metric(year, "收盘价").apply(pd.to_numeric, errors="coerce")
        rating = clean_rating(read_year_metric(year, "债项评级"))
        status = read_year_metric(year, "交易状态").astype("string").apply(lambda col: col.str.strip())
        if close.empty or rating.empty or status.empty:
            continue
        dates = close.index.intersection(rating.index).intersection(status.index)
        codes = close.columns.intersection(rating.columns).intersection(status.columns)
        close = close.loc[dates, codes]
        rating = rating.loc[dates, codes]
        status = status.loc[dates, codes]
        valid = status.eq("交易") & close.notna()

        rows = []
        for code in codes:
            prices = close[code].where(valid[code]).dropna()
            if prices.empty:
                continue
            annual_return = float(prices.iloc[-1] / prices.iloc[0] - 1) if len(prices) >= 2 else np.nan
            rows.append(
                {
                    "年份": year,
                    "转债代码": code,
                    "转债名称": name.get(code, pd.NA),
                    "评级": first_valid_rating(rating[code], valid[code]),
                    "发行规模": issue_size.get(code, np.nan),
                    "申万行业": industry.get(code, pd.NA),
                    "首个正常交易日": prices.index[0],
                    "末个正常交易日": prices.index[-1],
                    "首日收盘价": float(prices.iloc[0]),
                    "末日收盘价": float(prices.iloc[-1]),
                    "收益率": annual_return,
                    "正常交易日数": int(len(prices)),
                }
            )
        if rows:
            year_details.append(pd.DataFrame(rows))

    detail = pd.concat(year_details, ignore_index=True)
    annual_stats = pd.concat(
        [make_group_stats(group, str(year)) for year, group in detail.groupby("年份", sort=True)],
        ignore_index=True,
    )

    first_observation = detail.sort_values(["首个正常交易日", "转债代码"]).drop_duplicates("转债代码")
    last_observation = detail.sort_values(["末个正常交易日", "转债代码"]).drop_duplicates("转债代码", keep="last")
    overall = first_observation[["转债代码", "转债名称", "评级", "发行规模", "申万行业", "首个正常交易日", "首日收盘价"]].copy()
    overall = overall.merge(
        last_observation[["转债代码", "末个正常交易日", "末日收盘价"]], on="转债代码", how="left"
    )
    overall["收益率"] = overall["末日收盘价"] / overall["首日收盘价"] - 1
    overall_stats = make_group_stats(overall, f"{START_YEAR}-至今")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    annual_stats.to_csv(OUTPUT_DIR / "年度评级统计.csv", index=False, encoding="utf-8-sig")
    overall_stats.to_csv(OUTPUT_DIR / "总体评级统计.csv", index=False, encoding="utf-8-sig")
    detail.to_csv(OUTPUT_DIR / "个券年度明细.csv", index=False, encoding="utf-8-sig")
    payload = {
        "start_year": START_YEAR,
        "end_year": current_year,
        "latest_date": str(pd.to_datetime(detail["末个正常交易日"]).max().date()),
        "overall_stats": overall_stats.replace({np.nan: None}).to_dict(orient="records"),
        "annual_stats": annual_stats.replace({np.nan: None}).to_dict(orient="records"),
        "detail_columns": list(detail.columns),
        "detail": detail.replace({np.nan: None, pd.NaT: None}).to_dict(orient="records"),
    }
    (OUTPUT_DIR / "stats.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    print("总体统计：")
    print(overall_stats.to_string(index=False))
    print("\n年度统计：")
    print(annual_stats.to_string(index=False))


if __name__ == "__main__":
    main()
