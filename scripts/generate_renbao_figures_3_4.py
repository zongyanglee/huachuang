from __future__ import annotations

import math
import warnings
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.ticker import FormatStrFormatter, MultipleLocator
import numpy as np
import pandas as pd
from scipy.optimize import OptimizeWarning, curve_fit


ROOT = Path(__file__).resolve().parents[2]
PARQUET_ROOT = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "人保周报20260810"
START_DATE = pd.Timestamp("2023-01-01")

RED = "#E6121B"
BLUE = "#0262BA"
GRAY = "#A6A6A6"
LIGHT_RED = "#E7B8B8"


def inverse_cubic(x, a, b, c, d):
    return a / np.power(x, 3) + b / np.power(x, 2) + c / x + d


def fit_premium_at_100(plain: pd.Series, premium: pd.Series, turnover: pd.Series) -> float:
    sample = pd.DataFrame(
        {
            "平价": pd.to_numeric(plain, errors="coerce"),
            "转股溢价率": pd.to_numeric(premium, errors="coerce"),
            "换手率": pd.to_numeric(turnover, errors="coerce"),
        }
    )
    sample = sample.replace(0, np.nan).dropna()
    sample = sample[
        sample["平价"].between(70, 130, inclusive="both")
        & (sample["换手率"] <= 50)
    ]
    if len(sample) < 8:
        return float("nan")

    low, high = sample["转股溢价率"].quantile([0.03, 0.97])
    sample = sample[
        (sample["转股溢价率"] > low) & (sample["转股溢价率"] < high)
    ]
    if len(sample) < 8:
        return float("nan")

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", OptimizeWarning)
            params, _ = curve_fit(
                inverse_cubic,
                sample["平价"].to_numpy(dtype=float),
                sample["转股溢价率"].to_numpy(dtype=float),
                maxfev=20000,
            )
        return float(inverse_cubic(100.0, *params))
    except Exception:
        return float("nan")


def load_daily_series() -> pd.DataFrame:
    master = pd.read_parquet(PARQUET_ROOT / "_special" / "总表.parquet").set_index("__row_id")
    listing = pd.to_datetime(master["上市日期"], errors="coerce")
    last_trade = pd.to_datetime(master["最后交易日"], errors="coerce")
    rows: list[dict[str, object]] = []

    files = sorted(PARQUET_ROOT.glob("20??/20????.parquet"))
    files = [f for f in files if int(f.stem[:4]) >= START_DATE.year]
    needed = ["收盘价", "平价", "转股溢价率", "换手率"]

    for file in files:
        raw = pd.read_parquet(file)
        date_cols = sorted(
            c for c in raw.columns if isinstance(c, str) and c[:4].isdigit()
        )
        date_cols = [c for c in date_cols if pd.Timestamp(c) >= START_DATE]
        if not date_cols:
            continue

        blocks: dict[str, pd.DataFrame] = {}
        for name in needed:
            block = raw.loc[
                raw["__sheet_name"].eq(name), ["__row_id", *date_cols]
            ].set_index("__row_id")
            blocks[name] = block.apply(pd.to_numeric, errors="coerce")

        all_ids = blocks["收盘价"].index
        listing_m = listing.reindex(all_ids)
        last_trade_m = last_trade.reindex(all_ids)

        for day in date_cols:
            dt = pd.Timestamp(day)
            close = blocks["收盘价"][day]
            valid_period = (
                (listing_m.isna() | (dt >= listing_m))
                & (last_trade_m.isna() | (dt <= last_trade_m))
                & close.notna()
            )
            ids = all_ids[valid_period.fillna(False)]
            rows.append(
                {
                    "日期": dt,
                    "转债价格中位数": float(close.reindex(ids).median()),
                    "百元平价拟合溢价率": fit_premium_at_100(
                        blocks["平价"].loc[ids, day],
                        blocks["转股溢价率"].loc[ids, day],
                        blocks["换手率"].loc[ids, day],
                    ),
                }
            )

    result = pd.DataFrame(rows).sort_values("日期")
    return result.drop_duplicates("日期", keep="last").reset_index(drop=True)


def two_month_ticks(start: pd.Timestamp, end: pd.Timestamp) -> list[pd.Timestamp]:
    tick = pd.Timestamp(start.year, start.month, 3)
    if tick < start:
        tick += pd.DateOffset(months=2)
    ticks: list[pd.Timestamp] = []
    while tick <= end:
        ticks.append(tick)
        tick += pd.DateOffset(months=2)
    return ticks


def format_date_tick(value, _position=None) -> str:
    dt = mdates.num2date(value)
    return f"{dt.year}/{dt.month}/{dt.day}"


def nice_ylim(values: pd.Series, *, step: float, zero_floor: bool) -> tuple[float, float]:
    finite = pd.to_numeric(values, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    low = float(finite.min())
    high = float(finite.max())
    pad = max(step * 0.25, (high - low) * 0.05)
    lower = math.floor((low - pad) / step) * step
    upper = math.ceil((high + pad) / step) * step
    if zero_floor:
        lower = max(0.0, lower)
    if lower == upper:
        upper += step
    return lower, upper


def plot_series(
    data: pd.DataFrame,
    column: str,
    output_name: str,
    legend_name: str,
    *,
    y_step: float,
    zero_floor: bool,
) -> dict[str, float]:
    series = data.set_index("日期")[column].dropna().sort_index()
    quantiles = series.quantile([0.25, 0.50, 0.75])

    plt.rcParams.update(
        {
            "font.family": ["SimSun", "Times New Roman"],
            "axes.unicode_minus": False,
            "font.size": 20,
        }
    )
    fig, ax = plt.subplots(figsize=(14.25, 8.535), dpi=200)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    ax.plot(series.index, series.values, color=RED, linewidth=3.8, label=legend_name, zorder=4)
    for q, color, label in [
        (0.25, BLUE, "25%"),
        (0.50, GRAY, "50%"),
        (0.75, LIGHT_RED, "75%"),
    ]:
        ax.axhline(
            float(quantiles.loc[q]),
            color=color,
            linewidth=3.5,
            linestyle=(0, (4.5, 4.5)),
            dash_capstyle="round",
            label=label,
            zorder=2,
        )

    start, end = series.index.min(), series.index.max()
    ax.set_xlim(start, end)
    ticks = two_month_ticks(start, end)
    ax.set_xticks(ticks)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(format_date_tick))
    plt.setp(ax.get_xticklabels(), rotation=48, ha="right", rotation_mode="anchor")

    all_y = pd.concat([series, pd.Series(quantiles.values)])
    ax.set_ylim(*nice_ylim(all_y, step=y_step, zero_floor=zero_floor))
    ax.yaxis.set_major_locator(MultipleLocator(y_step))
    ax.yaxis.set_major_formatter(FormatStrFormatter("%.2f"))

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for side in ["left", "bottom"]:
        ax.spines[side].set_linewidth(1.3)
        ax.spines[side].set_color("black")
    ax.tick_params(axis="both", which="major", labelsize=20, width=1.2, length=6, pad=8)
    ax.grid(False)

    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.25),
        ncol=4,
        frameon=False,
        fontsize=21,
        handlelength=4.5,
        columnspacing=2.0,
        handletextpad=0.6,
    )
    fig.subplots_adjust(left=0.10, right=0.985, top=0.965, bottom=0.31)
    fig.savefig(OUTPUT_DIR / output_name, dpi=200, facecolor="white")
    plt.close(fig)

    return {
        "最新值": float(series.iloc[-1]),
        "25%分位数": float(quantiles.loc[0.25]),
        "50%分位数": float(quantiles.loc[0.50]),
        "75%分位数": float(quantiles.loc[0.75]),
    }


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    data = load_daily_series()

    summary = {
        "图3": plot_series(
            data,
            "转债价格中位数",
            "图3_转债价格中位数及历史分位数.png",
            "转债价格中位数",
            y_step=10.0,
            zero_floor=False,
        ),
        "图4": plot_series(
            data,
            "百元平价拟合溢价率",
            "图4_转债百元拟合溢价率及历史分位数.png",
            "百元平价拟合溢价率",
            y_step=5.0,
            zero_floor=True,
        ),
    }

    print(f"日期：{data['日期'].min().date()} 至 {data['日期'].max().date()}，共{len(data)}个交易日")
    for key, stats in summary.items():
        print(key, "，".join(f"{name}={value:.4f}" for name, value in stats.items()))


if __name__ == "__main__":
    main()
