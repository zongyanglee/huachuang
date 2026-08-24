"""Generate the convertible-bond and equity state timeline figure."""

from __future__ import annotations

from bisect import bisect_left, bisect_right, insort_right
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "outputs" / "转债与权益状态时序图_20260807"
OUTPUT_PATH = OUTPUT_DIR / "转债与权益状态时序图.png"

SHEETS = ["收盘价", "正股收盘价", "指数"]
START = pd.Timestamp("2015-01-01")
DISPLAY_START = pd.Timestamp("2019-01-01")


def chinese_font() -> fm.FontProperties:
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return fm.FontProperties(fname=candidate)
    return fm.FontProperties()


FONT = chinese_font()


def month_files() -> list[Path]:
    return sorted(
        path
        for path in DATA_DIR.glob("20*/*.parquet")
        if path.parent.name.isdigit() and path.name[:6].isdigit()
    )


def load_daily_panel() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Read the three needed sheets from monthly parquet files and reshape dates."""
    close_parts: list[pd.DataFrame] = []
    stock_parts: list[pd.DataFrame] = []
    index_parts: list[pd.DataFrame] = []

    for path in month_files():
        raw = pd.read_parquet(path, filters=[("__sheet_name", "in", SHEETS)])
        date_cols = [column for column in raw.columns if len(str(column)) == 10 and str(column)[4] == "-"]
        long = raw.melt(
            id_vars=["__sheet_name", "__row_id"],
            value_vars=date_cols,
            var_name="date",
            value_name="value",
        )
        long["date"] = pd.to_datetime(long["date"])
        long["value"] = pd.to_numeric(long["value"], errors="coerce")
        long = long[long["date"] >= START]

        for sheet, destination in [
            ("收盘价", close_parts),
            ("正股收盘价", stock_parts),
            ("指数", index_parts),
        ]:
            subset = long.loc[long["__sheet_name"].eq(sheet), ["date", "__row_id", "value"]]
            if not subset.empty:
                destination.append(subset)

    def combine(parts: list[pd.DataFrame]) -> pd.DataFrame:
        data = pd.concat(parts, ignore_index=True)
        return data.pivot(index="date", columns="__row_id", values="value").sort_index()

    return combine(close_parts), combine(stock_parts), combine(index_parts)


def complete_windows(mask: pd.Series) -> list[tuple[pd.Timestamp, pd.Timestamp]]:
    """Turn a Boolean daily state into continuous date spans for axvspan."""
    state_change = mask.ne(mask.shift()).cumsum()
    windows: list[tuple[pd.Timestamp, pd.Timestamp]] = []
    for _, group in mask.groupby(state_change):
        if bool(group.iloc[0]):
            windows.append((group.index[0], group.index[-1]))
    return windows


def historical_percentile(series: pd.Series) -> pd.Series:
    """As-of percentile rank using only observations available through each date."""
    observed: list[float] = []
    result = pd.Series(index=series.index, dtype="float64")
    for date, value in series.items():
        if pd.isna(value):
            continue
        numeric_value = float(value)
        left = bisect_left(observed, numeric_value)
        right = bisect_right(observed, numeric_value)
        count = len(observed) + 1
        # Average rank with the current observation included, matching rank(pct=True).
        average_rank = ((left + 1) + (right + 1)) / 2
        result.loc[date] = average_rank / count * 100
        insort_right(observed, numeric_value)
    return result


def make_figure() -> None:
    close, stock_close, indices = load_daily_panel()

    cb_median_price = close.median(axis=1, skipna=True).rename("转债价格中位数")
    stock_return_120 = stock_close.pct_change(120, fill_method=None).mean(axis=1, skipna=True).rename("正股等权120日收益")
    price_pct = historical_percentile(cb_median_price).rename("转债价格中位数历史分位")
    stock_return_pct = historical_percentile(stock_return_120).rename("正股120日收益历史分位")

    state = pd.concat([price_pct, stock_return_pct], axis=1).dropna()
    state = state.loc[state.index >= DISPLAY_START]
    target = state["转债价格中位数历史分位"].le(40) & state["正股120日收益历史分位"].ge(80)

    index_panel = indices.reindex(columns=["转债指数", "正股等权指数"]).dropna(how="all")
    index_panel = index_panel.loc[index_panel.index >= DISPLAY_START].ffill().dropna()
    base = index_panel.iloc[0]
    index_normalized = index_panel.div(base).mul(100)

    plt.rcParams.update({
        "axes.unicode_minus": False,
        "font.family": FONT.get_name(),
        "font.size": 12,
        "axes.titleweight": "bold",
        "axes.labelcolor": "#24292f",
        "xtick.color": "#4b5563",
        "ytick.color": "#4b5563",
    })
    fig, (ax_top, ax_bottom) = plt.subplots(
        2,
        1,
        figsize=(16, 9),
        dpi=180,
        sharex=True,
        gridspec_kw={"height_ratios": [1.04, 0.96], "hspace": 0.18},
    )
    fig.patch.set_facecolor("white")

    shade = "#f5d7b2"
    for start, end in complete_windows(target):
        for axis in (ax_top, ax_bottom):
            axis.axvspan(start, end + pd.Timedelta(days=1), color=shade, alpha=0.42, lw=0, zorder=0)

    price_line = ax_top.plot(
        state.index, state["转债价格中位数历史分位"], color="#1f5a99", lw=1.65, label="转债价格中位数历史分位"
    )[0]
    stock_line = ax_top.plot(
        state.index, state["正股120日收益历史分位"], color="#c65432", lw=1.65, label="正股120日收益分位"
    )[0]
    threshold_40 = ax_top.axhline(40, color="#1f5a99", lw=1.1, ls="--", alpha=0.8)
    threshold_80 = ax_top.axhline(80, color="#c65432", lw=1.1, ls="--", alpha=0.8)
    ax_top.text(state.index[0], 42.5, "40%", color="#1f5a99", va="bottom", ha="left", fontproperties=FONT, fontsize=10)
    ax_top.text(state.index[0], 82.5, "80%", color="#c65432", va="bottom", ha="left", fontproperties=FONT, fontsize=10)
    ax_top.set_ylim(-2, 102)
    ax_top.set_yticks([0, 20, 40, 60, 80, 100])
    ax_top.set_yticklabels([f"{value}%" for value in [0, 20, 40, 60, 80, 100]], fontproperties=FONT)
    ax_top.set_ylabel("历史分位", fontproperties=FONT)
    ax_top.set_title("转债与权益状态时序图", loc="left", fontsize=18, pad=12, fontproperties=FONT)
    ax_top.text(
        0,
        1.01,
        "阴影：转债价格中位数分位 ≤ 40% 且正股120日收益分位 ≥ 80%",
        transform=ax_top.transAxes,
        color="#6b7280",
        fontsize=11,
        fontproperties=FONT,
    )
    ax_top.legend(
        [price_line, stock_line],
        ["转债价格中位数历史分位", "正股120日收益分位"],
        loc="upper right",
        frameon=False,
        prop=FONT,
        ncol=2,
    )

    cb_index = ax_bottom.plot(index_normalized.index, index_normalized["转债指数"], color="#1f5a99", lw=1.8, label="转债指数")[0]
    stock_index = ax_bottom.plot(index_normalized.index, index_normalized["正股等权指数"], color="#c65432", lw=1.8, label="正股等权指数")[0]
    ax_bottom.set_ylabel("净值（2019-01-02=100）", fontproperties=FONT)
    ax_bottom.legend([cb_index, stock_index], ["转债指数", "正股等权指数"], loc="upper right", frameon=False, prop=FONT, ncol=2)

    # Representative windows selected by the five-quantile matrix itself,
    # rather than by ex-post market performance or macro policy announcements.
    event_specs = [
        ("2019-02-22", "2019.02-08\n低价 × 强正股", 0.98, "left"),
        ("2020-05-14", "2020.05-07\n低价 × 强正股", 0.78, "left"),
        ("2020-12-11", "2020.12\n低价 × 强正股", 0.98, "left"),
    ]
    y_min, y_max = ax_bottom.get_ylim()
    y_span = y_max - y_min
    for date_text, label, y_fraction, alignment in event_specs:
        event_date = pd.Timestamp(date_text)
        ax_bottom.axvline(event_date, color="#7b8794", lw=0.9, ls=":", zorder=1)
        ax_bottom.annotate(
            label,
            xy=(event_date, y_min + y_span * 0.06),
            xytext=(event_date, y_min + y_span * y_fraction),
            ha=alignment,
            va="top",
            color="#465362",
            fontsize=10.5,
            fontproperties=FONT,
            arrowprops={"arrowstyle": "-", "color": "#7b8794", "lw": 0.8},
        )

    for axis in (ax_top, ax_bottom):
        axis.grid(axis="y", color="#d9dee5", lw=0.7, alpha=0.85)
        axis.spines[["top", "right"]].set_visible(False)
        axis.spines[["left", "bottom"]].set_color("#aab4c0")
        axis.margins(x=0)

    ax_bottom.xaxis.set_major_locator(mdates.YearLocator())
    ax_bottom.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax_bottom.set_xlim(DISPLAY_START - pd.Timedelta(days=16), index_normalized.index[-1] + pd.Timedelta(days=16))
    for label in ax_bottom.get_xticklabels():
        label.set_fontproperties(FONT)
    for label in ax_bottom.get_yticklabels():
        label.set_fontproperties(FONT)

    fig.text(
        0.01,
        0.015,
        "数据来源：转债个券历史序列；转债价格为存续样本收盘价中位数，正股收益为对应正股120个交易日收益的等权平均；历史分位仅使用2015年起截至当日的数据。",
        color="#6b7280",
        fontsize=9.5,
        fontproperties=FONT,
    )
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(OUTPUT_PATH, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    make_figure()
    print(OUTPUT_PATH)
