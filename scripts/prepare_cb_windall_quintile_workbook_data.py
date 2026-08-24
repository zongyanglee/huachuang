"""Prepare auditable daily data for the CB/Wind All A quintile workbook."""

from __future__ import annotations

import json
from bisect import bisect_left, bisect_right, insort_right
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "outputs" / "转债价格与万得全A五分位矩阵_20260807"
OUTPUT_JSON = OUTPUT_DIR / "matrix_data.json"
STAT_START = pd.Timestamp("2019-01-01")


def historical_percentile(series: pd.Series) -> pd.Series:
    observed: list[float] = []
    result = pd.Series(index=series.index, dtype="float64")
    for date, value in series.items():
        if pd.isna(value):
            continue
        numeric_value = float(value)
        left = bisect_left(observed, numeric_value)
        right = bisect_right(observed, numeric_value)
        count = len(observed) + 1
        result.loc[date] = (((left + 1) + (right + 1)) / 2) / count * 100
        insort_right(observed, numeric_value)
    return result


def load_daily_panel() -> tuple[pd.DataFrame, pd.DataFrame]:
    close_parts: list[pd.DataFrame] = []
    index_parts: list[pd.DataFrame] = []
    for path in sorted(DATA_DIR.glob("20*/*.parquet")):
        raw = pd.read_parquet(path, filters=[("__sheet_name", "in", ["收盘价", "指数"])])
        date_cols = [column for column in raw.columns if len(str(column)) == 10 and str(column)[4] == "-"]
        long = raw.melt(
            id_vars=["__sheet_name", "__row_id"], value_vars=date_cols, var_name="date", value_name="value"
        )
        long["date"] = pd.to_datetime(long["date"])
        long["value"] = pd.to_numeric(long["value"], errors="coerce")
        close_parts.append(long.loc[long["__sheet_name"].eq("收盘价"), ["date", "__row_id", "value"]])
        index_parts.append(long.loc[long["__sheet_name"].eq("指数"), ["date", "__row_id", "value"]])

    def combine(parts: list[pd.DataFrame]) -> pd.DataFrame:
        data = pd.concat(parts, ignore_index=True)
        return data.pivot(index="date", columns="__row_id", values="value").sort_index()

    return combine(close_parts), combine(index_parts)


def optional_number(value: object) -> float | int | None:
    if pd.isna(value):
        return None
    if isinstance(value, (np.integer, int)):
        return int(value)
    return float(value)


def quintile(percentile: pd.Series) -> pd.Series:
    return np.ceil(percentile.mul(5)).clip(1, 5).astype("Int64")


def main() -> None:
    close, indices = load_daily_panel()
    cb_price_median = close.median(axis=1, skipna=True).rename("cb_price_median")
    wind_all_a = indices.iloc[:, 0].rename("wind_all_a")
    cb_index = indices.iloc[:, 4].rename("cb_index")

    wind_return_120 = wind_all_a.pct_change(120, fill_method=None).rename("wind_return_120")
    cb_price_pct = historical_percentile(cb_price_median).div(100).rename("cb_price_pct")
    wind_return_pct = historical_percentile(wind_return_120).div(100).rename("wind_return_pct")

    frame = pd.concat(
        [cb_price_median, cb_price_pct, wind_all_a, wind_return_120, wind_return_pct, cb_index], axis=1
    ).sort_index()
    frame["cb_price_q"] = quintile(frame["cb_price_pct"])
    frame["wind_return_q"] = quintile(frame["wind_return_pct"])
    frame["future20_cb_index"] = frame["cb_index"].shift(-20)
    frame["future20_cb_return"] = frame["future20_cb_index"].div(frame["cb_index"]).sub(1)
    frame["positive_flag"] = frame["future20_cb_return"].gt(0).where(frame["future20_cb_return"].notna())
    frame["in_sample"] = (
        (frame.index >= STAT_START)
        & frame["cb_price_q"].notna()
        & frame["wind_return_q"].notna()
        & frame["future20_cb_return"].notna()
    )

    sample = frame.loc[frame["in_sample"]].copy()
    expected: dict[str, list[list[float | int | None]]] = {}
    for metric, column, agg in [
        ("win_rate", "positive_flag", "mean"),
        ("avg_return", "future20_cb_return", "mean"),
        ("count", "future20_cb_return", "count"),
    ]:
        table = sample.pivot_table(
            index="cb_price_q", columns="wind_return_q", values=column, aggfunc=agg, observed=True
        ).reindex(index=[1, 2, 3, 4, 5], columns=[1, 2, 3, 4, 5])
        expected[metric] = [[optional_number(value) for value in row] for row in table.to_numpy()]

    records: list[list[object]] = []
    for date, row in frame.iterrows():
        records.append(
            [
                date.strftime("%Y-%m-%d"),
                optional_number(row["cb_price_median"]),
                optional_number(row["cb_price_pct"]),
                optional_number(row["cb_price_q"]),
                optional_number(row["wind_all_a"]),
                optional_number(row["wind_return_120"]),
                optional_number(row["wind_return_pct"]),
                optional_number(row["wind_return_q"]),
                optional_number(row["cb_index"]),
                optional_number(row["future20_cb_index"]),
                optional_number(row["future20_cb_return"]),
                None if pd.isna(row["positive_flag"]) else int(bool(row["positive_flag"])),
                int(bool(row["in_sample"])),
            ]
        )

    payload = {
        "metadata": {
            "history_start": frame.index.min().strftime("%Y-%m-%d"),
            "stat_start": sample.index.min().strftime("%Y-%m-%d"),
            "stat_end": sample.index.max().strftime("%Y-%m-%d"),
            "data_end": frame.index.max().strftime("%Y-%m-%d"),
            "rows": len(frame),
            "sample_rows": len(sample),
        },
        "headers": [
            "日期",
            "转债价格中位数",
            "转债价格中位数截至当日历史分位",
            "转债价格分位组",
            "万得全A指数",
            "万得全A过去120日收益",
            "万得全A120日收益截至当日历史分位",
            "万得全A分位组",
            "转债指数",
            "未来20日转债指数",
            "未来20日转债指数收益率",
            "未来20日正收益标记",
            "纳入2019年以来统计",
        ],
        "records": records,
        "expected": expected,
    }
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    print(OUTPUT_JSON)


if __name__ == "__main__":
    main()
