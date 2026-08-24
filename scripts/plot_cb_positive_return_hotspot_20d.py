"""Plot the five-quintile hot-spot signal against forward 20-day CB returns."""

from __future__ import annotations

from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from plot_cb_equity_state_timeline import historical_percentile, load_daily_panel


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "outputs" / "五分位矩阵结论_转债未来20日正收益_20260807"
OUTPUT_PATH = OUTPUT_DIR / "低价转债强权益动量_未来20日正收益验证.png"
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
    close, stock_close, indices = load_daily_panel()
    cb_price_median = close.median(axis=1, skipna=True).rename("cb_price_median")
    stock_return_120 = stock_close.pct_change(120, fill_method=None).mean(axis=1, skipna=True).rename(
        "stock_return_120"
    )
    cb_price_pct = historical_percentile(cb_price_median).rename("cb_price_pct")
    stock_return_pct = historical_percentile(stock_return_120).rename("stock_return_pct")
    daily = pd.concat(
        [
            cb_price_median,
            stock_return_120,
            cb_price_pct,
            stock_return_pct,
            indices.iloc[:, 4].rename("cb_index"),
        ],
        axis=1,
    ).dropna()
    daily["cb_return_20"] = daily["cb_index"].shift(-20).div(daily["cb_index"]).sub(1).mul(100)
    daily = daily.loc[daily.index >= START].copy()

    weekly = daily.groupby(daily.index.to_period("W-FRI")).tail(1).copy()
    weekly["signal"] = weekly["cb_price_pct"].le(40) & weekly["stock_return_pct"].ge(60)
    return daily, weekly


def draw() -> None:
    daily, weekly = prepare_data()
    signal = weekly.loc[weekly["signal"] & weekly["cb_return_20"].notna()].copy()
    cb_normalized = daily["cb_index"].div(daily["cb_index"].iloc[0]).mul(100)
    forward_return = daily["cb_return_20"]
    values = forward_return.to_numpy(dtype=float, na_value=np.nan)
    finite = np.isfinite(values)

    win_rate = signal["cb_return_20"].gt(0).mean() * 100
    average_return = signal["cb_return_20"].mean()

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
    fig, (axis_index, axis_return) = plt.subplots(
        2,
        1,
        figsize=(16, 9.2),
        dpi=180,
        sharex=True,
        gridspec_kw={"height_ratios": [1.08, 0.92], "hspace": 0.12},
    )
    fig.patch.set_facecolor("white")
    fig.suptitle(
        "低价转债 × 强权益动量：未来20日正收益验证",
        x=0.075,
        y=0.98,
        ha="left",
        fontsize=18,
        fontproperties=FONT,
    )
    fig.text(
        0.075,
        0.94,
        "周度信号：转债价格中位数历史分位≤40%，正股等权120日收益历史分位≥60%（均按截至当日历史计算）",
        color="#687383",
        fontsize=10.5,
        fontproperties=FONT,
    )

    for date in signal.index:
        for axis in (axis_index, axis_return):
            axis.axvspan(
                date - pd.Timedelta(days=4),
                date + pd.Timedelta(days=3),
                color="#f3c66b",
                alpha=0.28,
                lw=0,
                zorder=0,
            )

    axis_index.plot(cb_normalized.index, cb_normalized, color="#1f5a99", lw=1.9, label="转债指数")
    axis_index.set_ylabel("净值（2019-01-02=100）", fontproperties=FONT)
    axis_index.legend(loc="upper left", frameon=False, prop=FONT)
    axis_index.text(
        0.985,
        0.94,
        f"信号周 {len(signal)} 期｜正收益胜率 {win_rate:.1f}%｜平均收益 {average_return:+.2f}%",
        transform=axis_index.transAxes,
        ha="right",
        va="top",
        fontsize=10.2,
        color="#4c5866",
        fontproperties=FONT,
        bbox={"boxstyle": "round,pad=0.35", "facecolor": "white", "edgecolor": "#d9dee5", "alpha": 0.92},
    )

    axis_return.plot(forward_return.index, values, color="#44566a", lw=1.35, label="转债未来20日收益率")
    axis_return.axhline(0, color="#6c7885", lw=1.0)
    axis_return.fill_between(
        forward_return.index,
        0,
        values,
        where=finite & (values >= 0),
        color="#cf513d",
        alpha=0.23,
    )
    axis_return.fill_between(
        forward_return.index,
        0,
        values,
        where=finite & (values < 0),
        color="#3d79aa",
        alpha=0.18,
    )
    axis_return.set_ylabel("未来20日收益率（%）", fontproperties=FONT)
    axis_return.legend(loc="upper right", frameon=False, prop=FONT)

    y_min, y_max = axis_return.get_ylim()
    span = y_max - y_min
    annotations = [
        (
            "2019-03-01",
            "2019-05-20",
            "2019.02—03  宽信用预期推动修复\n信号期未来20日均为正\n平均收益 +4.5%",
            0.94,
        ),
        (
            "2019-12-06",
            "2019-12-06",
            "2019.11—12  风险偏好回升\n信号期未来20日均为正\n平均收益 +4.1%",
            0.69,
        ),
        (
            "2020-06-12",
            "2020-07-10",
            "2020.05—06  复苏交易加速\n信号期未来20日均为正\n平均收益 +5.1%",
            0.94,
        ),
        (
            "2024-10-25",
            "2024-08-20",
            "2024.09—11  政策组合改善预期\n信号期未来20日均为正\n平均收益 +2.5%",
            0.82,
        ),
    ]
    for date_text, text_date, label, fraction in annotations:
        date = pd.Timestamp(date_text)
        nearest = forward_return.index[forward_return.index.get_indexer([date], method="nearest")[0]]
        y_value = float(forward_return.loc[nearest]) if pd.notna(forward_return.loc[nearest]) else 0.0
        axis_return.annotate(
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

    for axis in (axis_index, axis_return):
        axis.grid(axis="y", color="#d9dee5", lw=0.7, alpha=0.88)
        axis.spines[["top", "right"]].set_visible(False)
        axis.spines[["left", "bottom"]].set_color("#aab4c0")
        axis.margins(x=0)
        for label in axis.get_yticklabels():
            label.set_fontproperties(FONT)

    axis_return.xaxis.set_major_locator(mdates.YearLocator())
    axis_return.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    axis_return.set_xlim(START - pd.Timedelta(days=16), daily.index[-1] + pd.Timedelta(days=16))
    for label in axis_return.get_xticklabels():
        label.set_fontproperties(FONT)

    fig.text(
        0.075,
        0.02,
        "注：黄色阴影为符合五分位条件的周度观察；下图为各交易日对应的转债指数未来20日收益率，红色表示正收益、蓝色表示负收益。历史分位仅使用2015年起截至当日数据。",
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
