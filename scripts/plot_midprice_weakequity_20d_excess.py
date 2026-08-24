"""Plot the corrected as-of percentile signal and forward 20-day excess return."""

from __future__ import annotations

from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from plot_cb_equity_state_timeline import historical_percentile, load_daily_panel


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "outputs" / "转债中位价格权益弱势_未来20日超额_20260807"
OUTPUT_PATH = OUTPUT_DIR / "转债相对万得全A_未来20日超额状态.png"
START = pd.Timestamp("2019-01-01")


def chinese_font() -> fm.FontProperties:
    for candidate in [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    ]:
        if Path(candidate).exists():
            return fm.FontProperties(fname=candidate)
    return fm.FontProperties()


FONT = chinese_font()


def prepare_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    close, _, indices = load_daily_panel()
    daily = pd.concat(
        [
            close.median(axis=1, skipna=True).rename("cb_price_median"),
            indices.iloc[:, 2].rename("stock_equal_index"),
            indices.iloc[:, 4].rename("cb_index"),
            indices.iloc[:, 0].rename("wind_all_a"),
        ],
        axis=1,
    ).dropna()

    # Both percentiles are expanding/as-of percentiles; no data after the date are used.
    daily["cb_price_pct"] = historical_percentile(daily["cb_price_median"])
    daily["stock_return_120"] = daily["stock_equal_index"].pct_change(120, fill_method=None)
    daily["stock_return_pct"] = historical_percentile(daily["stock_return_120"])
    daily["cb_return_20"] = daily["cb_index"].shift(-20).div(daily["cb_index"]).sub(1)
    daily["wind_return_20"] = daily["wind_all_a"].shift(-20).div(daily["wind_all_a"]).sub(1)
    daily["excess_20"] = daily["cb_return_20"].sub(daily["wind_return_20"]).mul(100)
    daily = daily.loc[daily.index >= START].copy()

    weekly = daily.groupby(daily.index.to_period("W-FRI")).tail(1).copy()
    weekly["signal"] = weekly["cb_price_pct"].between(40, 60, inclusive="both") & weekly[
        "stock_return_pct"
    ].between(20, 40, inclusive="both")
    return daily, weekly


def draw() -> None:
    daily, weekly = prepare_data()
    signal = weekly.loc[weekly["signal"]].copy()
    normalized = daily[["cb_index", "wind_all_a"]].div(daily[["cb_index", "wind_all_a"]].iloc[0]).mul(100)
    excess = daily["excess_20"]
    excess_values = excess.to_numpy(dtype=float, na_value=np.nan)
    finite = np.isfinite(excess_values)

    signal_valid = signal["excess_20"].dropna()
    excess_win = signal_valid.gt(0).mean() * 100
    average_excess = signal_valid.mean()

    plt.rcParams.update(
        {
            "font.family": FONT.get_name(),
            "axes.unicode_minus": False,
            "font.size": 11.5,
            "axes.labelcolor": "#2f3945",
            "xtick.color": "#53606d",
            "ytick.color": "#53606d",
        }
    )
    fig, (axis_index, axis_excess) = plt.subplots(
        2,
        1,
        figsize=(16, 9.2),
        dpi=180,
        sharex=True,
        gridspec_kw={"height_ratios": [1.1, 0.9], "hspace": 0.12},
    )
    fig.patch.set_facecolor("white")
    fig.suptitle(
        "转债相对万得全A超额状态：未来20日收益差验证",
        x=0.075,
        y=0.98,
        ha="left",
        fontsize=18,
        fontproperties=FONT,
    )
    fig.text(
        0.075,
        0.94,
        "周度信号：转债价格中位数历史分位40%—60%，正股等权指数过去120日收益历史分位20%—40%（均按截至当日历史计算）",
        color="#687383",
        fontsize=10.5,
        fontproperties=FONT,
    )

    for date in signal.index:
        for axis in (axis_index, axis_excess):
            axis.axvspan(
                date - pd.Timedelta(days=4),
                date + pd.Timedelta(days=3),
                color="#f3c66b",
                alpha=0.28,
                lw=0,
                zorder=0,
            )

    axis_index.plot(normalized.index, normalized["cb_index"], color="#1f5a99", lw=1.8, label="转债指数")
    axis_index.plot(normalized.index, normalized["wind_all_a"], color="#c65432", lw=1.8, label="万得全A")
    axis_index.set_ylabel("净值（2019-01-02=100）", fontproperties=FONT)
    axis_index.legend(loc="upper left", frameon=False, prop=FONT, ncol=2)

    stats_text = f"信号周 {len(signal_valid)} 期｜超额胜率 {excess_win:.1f}%｜平均超额 {average_excess:+.2f}%"
    axis_index.text(
        0.985,
        0.94,
        stats_text,
        transform=axis_index.transAxes,
        ha="right",
        va="top",
        fontsize=10.2,
        color="#4c5866",
        fontproperties=FONT,
        bbox={"boxstyle": "round,pad=0.35", "facecolor": "white", "edgecolor": "#d9dee5", "alpha": 0.92},
    )

    axis_excess.plot(excess.index, excess_values, color="#44566a", lw=1.35, label="未来20日收益率差值")
    axis_excess.axhline(0, color="#6c7885", lw=1.0)
    axis_excess.fill_between(
        excess.index, 0, excess_values, where=finite & (excess_values >= 0), color="#cf513d", alpha=0.23
    )
    axis_excess.fill_between(
        excess.index, 0, excess_values, where=finite & (excess_values < 0), color="#3d79aa", alpha=0.18
    )
    axis_excess.set_ylabel("未来20日：转债－万得全A（%）", fontproperties=FONT)
    axis_excess.legend(loc="upper left", frameon=False, prop=FONT)

    y_min, y_max = axis_excess.get_ylim()
    span = y_max - y_min
    annotations = [
        (
            "2021-02-26",
            "2021-02-26",
            "2021.01—04  权益由上涨转震荡\n转债 +1.0%，万得全A -0.2%\n平均超额 +1.2%",
            0.90,
        ),
        (
            "2022-12-30",
            "2022-12-30",
            "2022.12  复苏交易快速升温\n转债 +4.7%，万得全A +8.3%\n平均超额 -3.6%",
            0.22,
        ),
        (
            "2023-12-22",
            "2023-10-25",
            "2023.12—2024.01  权益下探\n转债 -1.8%，万得全A -7.3%\n平均超额 +5.5%",
            0.92,
        ),
        (
            "2024-04-30",
            "2024-07-05",
            "2024.03—05  修复后转入震荡\n转债 +0.1%，万得全A -2.3%\n平均超额 +2.4%",
            0.56,
        ),
    ]
    for date_text, text_date, label, fraction in annotations:
        date = pd.Timestamp(date_text)
        nearest = excess.index[excess.index.get_indexer([date], method="nearest")[0]]
        y_value = float(excess.loc[nearest]) if pd.notna(excess.loc[nearest]) else 0.0
        axis_excess.annotate(
            label,
            xy=(nearest, y_value),
            xytext=(pd.Timestamp(text_date), y_min + span * fraction),
            ha="center",
            va="top",
            fontsize=9.0,
            color="#4c5866",
            fontproperties=FONT,
            arrowprops={"arrowstyle": "-", "color": "#8b96a3", "lw": 0.8},
        )

    for axis in (axis_index, axis_excess):
        axis.grid(axis="y", color="#d9dee5", lw=0.7, alpha=0.88)
        axis.spines[["top", "right"]].set_visible(False)
        axis.spines[["left", "bottom"]].set_color("#aab4c0")
        axis.margins(x=0)
        for label in axis.get_yticklabels():
            label.set_fontproperties(FONT)

    axis_excess.xaxis.set_major_locator(mdates.YearLocator())
    axis_excess.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    axis_excess.set_xlim(START - pd.Timedelta(days=16), daily.index[-1] + pd.Timedelta(days=16))
    for label in axis_excess.get_xticklabels():
        label.set_fontproperties(FONT)

    fig.text(
        0.075,
        0.02,
        "注：黄色阴影为符合条件的周度观察；未来20日收益率差值=转债指数未来20日收益率－万得全A未来20日收益率，差值>0表示转债取得超额。历史分位仅使用2015年起截至当日数据。",
        color="#687383",
        fontsize=9.2,
        fontproperties=FONT,
    )
    fig.subplots_adjust(left=0.075, right=0.985, top=0.89, bottom=0.085)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(OUTPUT_PATH, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(OUTPUT_PATH)


if __name__ == "__main__":
    draw()
