#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Analyze convertible-bond performance by credit rating in summer 2024."""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "outputs" / "rating_nav_202406_202408"
START = pd.Timestamp("2024-06-03")
END = pd.Timestamp("2024-08-30")
RATINGS = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-"]
# Exclude listing-day returns; a new bond enters on its first subsequent
# normal-trading day.
VALID_STATUS = {"交易"}
EXCLUDED_CODES = {"113589.SH", "128082.SZ"}  # 天创转债、华锋转债
LEVERAGE_SAMPLES = {
    "113519.SH": {
        "发行人": "长久物流",
        "正股代码": "603569.SH",
        "2023资产负债率": 0.4318989174,
    },
    "113524.SH": {
        "发行人": "奇精机械",
        "正股代码": "603677.SH",
        "2023资产负债率": 0.460445982208,
    },
    "113542.SH": {
        "发行人": "好莱客",
        "正股代码": "603898.SH",
        "2023资产负债率": 0.319446369445,
    },
}


def read_metric(metric: str) -> pd.DataFrame:
    parts: list[pd.DataFrame] = []
    for month in ("202406", "202407", "202408"):
        frame = pd.read_parquet(DATA_ROOT / "2024" / f"{month}.parquet")
        sub = frame.loc[frame["__sheet_name"].eq(metric)].copy()
        date_cols = [c for c in sub.columns if c not in {"__sheet_name", "__row_id"}]
        parsed = pd.to_datetime(pd.Index(date_cols), errors="coerce")
        keep = [
            col
            for col, date in zip(date_cols, parsed)
            if pd.notna(date) and START <= pd.Timestamp(date) <= END
        ]
        wide = sub.set_index("__row_id")[keep].T
        wide.index = pd.to_datetime(wide.index)
        wide.columns = wide.columns.astype(str)
        parts.append(wide)
    result = pd.concat(parts).sort_index()
    return result.loc[~result.index.duplicated(keep="last")]


def clean_rating(frame: pd.DataFrame) -> pd.DataFrame:
    return (
        frame.astype("string")
        .replace({"0": pd.NA, "": pd.NA, "nan": pd.NA, "None": pd.NA})
        .apply(lambda col: col.str.strip().str.upper().str.replace("＋", "+").str.replace("－", "-"))
    )


def max_drawdown(series: pd.Series) -> float:
    series = series.dropna()
    if series.empty:
        return np.nan
    return float((series / series.cummax() - 1).min())


def first_valid(series: pd.Series) -> float:
    values = series.dropna()
    return float(values.iloc[0]) if not values.empty else np.nan


def last_valid(series: pd.Series) -> float:
    values = series.dropna()
    return float(values.iloc[-1]) if not values.empty else np.nan


def main() -> None:
    returns = read_metric("涨跌幅").apply(pd.to_numeric, errors="coerce") / 100
    close = read_metric("收盘价").apply(pd.to_numeric, errors="coerce")
    ratings = clean_rating(read_metric("债项评级"))
    status = read_metric("交易状态").astype("string").apply(lambda col: col.str.strip())
    balance = read_metric("余额").apply(pd.to_numeric, errors="coerce")
    stock_cap = read_metric("正股市值").apply(pd.to_numeric, errors="coerce")

    common_dates = returns.index.intersection(close.index).intersection(ratings.index).intersection(status.index)
    common_codes = (
        returns.columns.intersection(close.columns).intersection(ratings.columns).intersection(status.columns)
        .difference(EXCLUDED_CODES)
    )
    returns = returns.loc[common_dates, common_codes]
    close = close.loc[common_dates, common_codes]
    ratings = ratings.loc[common_dates, common_codes]
    status = status.loc[common_dates, common_codes]
    balance = balance.reindex(index=common_dates, columns=common_codes)
    stock_cap = stock_cap.reindex(index=common_dates, columns=common_codes)

    valid = status.isin(VALID_STATUS) & close.notna() & returns.notna()
    daily_return = pd.DataFrame(index=common_dates)
    holding_count = pd.DataFrame(index=common_dates)
    for rating in RATINGS:
        mask = ratings.eq(rating) & valid
        daily_return[rating] = returns.where(mask).mean(axis=1)
        holding_count[rating] = mask.sum(axis=1)

    daily_return = daily_return.dropna(how="all").fillna(0)
    nav = (1 + daily_return).cumprod()
    nav = nav.div(nav.iloc[0])

    summary_rows = []
    for rating in RATINGS:
        series = nav[rating]
        summary_rows.append(
            {
                "评级": rating,
                "期初净值": float(series.iloc[0]),
                "期末净值": float(series.iloc[-1]),
                "区间收益": float(series.iloc[-1] - 1),
                "最大回撤": max_drawdown(series),
                "日波动率": float(daily_return[rating].std(ddof=0)),
                "平均持仓数": float(holding_count[rating].mean()),
                "期初持仓数": int(holding_count[rating].iloc[0]),
                "期末持仓数": int(holding_count[rating].iloc[-1]),
            }
        )
    summary = pd.DataFrame(summary_rows)

    total = pd.read_parquet(DATA_ROOT / "_special" / "总表.parquet").set_index("__row_id")
    initial_rating = ratings.apply(lambda col: col.dropna().iloc[0] if col.notna().any() else pd.NA)
    detail_rows = []
    for code in common_codes:
        price = close[code].where(valid[code]).dropna()
        if len(price) < 20:
            continue
        first_date = price.index[0]
        last_date = price.index[-1]
        if first_date > START + pd.Timedelta(days=7) or last_date < END - pd.Timedelta(days=7):
            continue
        rating = initial_rating.get(code)
        if rating not in RATINGS:
            continue
        first_price = float(price.iloc[0])
        last_price = float(price.iloc[-1])
        issue_size = pd.to_numeric(total["发行规模"].get(code), errors="coerce") if code in total.index else np.nan
        name = total["转债名称"].get(code, "") if code in total.index else ""
        detail_rows.append(
            {
                "代码": code,
                "转债名称": name,
                "期初评级": rating,
                "首个有效日": first_date,
                "末个有效日": last_date,
                "期初收盘价": first_price,
                "期末收盘价": last_price,
                "区间收益": last_price / first_price - 1,
                "最大回撤": max_drawdown(price),
                "期初余额": first_valid(balance[code]),
                "发行规模": float(issue_size) if pd.notna(issue_size) else np.nan,
                "期初正股市值": first_valid(stock_cap[code]),
            }
        )
    detail = pd.DataFrame(detail_rows)

    high_worst = (
        detail.loc[detail["期初评级"].isin(["AAA", "AA+"])]
        .sort_values(["区间收益", "最大回撤"])
        .head(20)
        .reset_index(drop=True)
    )
    low_best = (
        detail.loc[detail["期初评级"].isin(["AA", "AA-", "A+", "A", "A-"])]
        .sort_values(["区间收益", "最大回撤"], ascending=[False, False])
        .head(20)
        .reset_index(drop=True)
    )
    defensive = detail.loc[detail["代码"].isin(LEVERAGE_SAMPLES)].copy()
    defensive["发行人"] = defensive["代码"].map(lambda code: LEVERAGE_SAMPLES[code]["发行人"])
    defensive["正股代码"] = defensive["代码"].map(lambda code: LEVERAGE_SAMPLES[code]["正股代码"])
    defensive["2023资产负债率"] = defensive["代码"].map(
        lambda code: LEVERAGE_SAMPLES[code]["2023资产负债率"]
    )
    defensive = defensive.sort_values("最大回撤", ascending=False).reset_index(drop=True)

    cross_rows = []
    for low_rating in ["AA", "AA-", "A+", "A", "A-"]:
        low = detail.loc[detail["期初评级"].eq(low_rating), "区间收益"].dropna()
        for high_rating in ["AAA", "AA+"]:
            high = detail.loc[detail["期初评级"].eq(high_rating), "区间收益"].dropna()
            if low.empty or high.empty:
                continue
            cross_rows.append(
                {
                    "低评级组": low_rating,
                    "高评级组": high_rating,
                    "低评级收益中位数": float(low.median()),
                    "高评级收益中位数": float(high.median()),
                    "中位数差": float(low.median() - high.median()),
                    "低评级跑赢高评级个券占比": float(
                        (low.to_numpy()[:, None] > high.to_numpy()[None, :]).mean()
                    ),
                    "低评级样本数": int(len(low)),
                    "高评级样本数": int(len(high)),
                }
            )
    cross = pd.DataFrame(cross_rows)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    nav.reset_index(names="日期").to_csv(OUTPUT_DIR / "nav.csv", index=False, encoding="utf-8-sig")
    daily_return.reset_index(names="日期").to_csv(
        OUTPUT_DIR / "daily_return.csv", index=False, encoding="utf-8-sig"
    )
    holding_count.reset_index(names="日期").to_csv(
        OUTPUT_DIR / "holding_count.csv", index=False, encoding="utf-8-sig"
    )
    summary.to_csv(OUTPUT_DIR / "summary.csv", index=False, encoding="utf-8-sig")
    detail.to_csv(OUTPUT_DIR / "detail.csv", index=False, encoding="utf-8-sig")
    high_worst.to_csv(OUTPUT_DIR / "high_worst.csv", index=False, encoding="utf-8-sig")
    low_best.to_csv(OUTPUT_DIR / "low_best.csv", index=False, encoding="utf-8-sig")
    defensive.to_csv(OUTPUT_DIR / "defensive.csv", index=False, encoding="utf-8-sig")
    cross.to_csv(OUTPUT_DIR / "cross.csv", index=False, encoding="utf-8-sig")

    report = {
        "start": str(common_dates.min().date()),
        "end": str(common_dates.max().date()),
        "trading_days": int(len(common_dates)),
        "nav": nav.reset_index(names="日期").to_dict(orient="records"),
        "summary": summary.to_dict(orient="records"),
        "detail": detail.to_dict(orient="records"),
        "high_worst": high_worst.head(10).to_dict(orient="records"),
        "low_best": low_best.head(10).to_dict(orient="records"),
        "defensive": defensive.to_dict(orient="records"),
        "cross": cross.to_dict(orient="records"),
    }
    (OUTPUT_DIR / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    print(summary.to_string(index=False))
    print("\n高评级跌幅靠前：")
    print(high_worst.head(10).to_string(index=False))
    print("\nAA及以下表现靠前：")
    print(low_best.head(10).to_string(index=False))
    print("\n跨评级比较：")
    print(cross.to_string(index=False))


if __name__ == "__main__":
    main()
