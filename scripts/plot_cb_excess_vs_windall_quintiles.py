"""Plot quintile matrices and historical validation for CB excess vs Wind All A."""

from __future__ import annotations

from pathlib import Path
import sys

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from plot_cb_equity_state_timeline import historical_percentile, load_daily_panel  # noqa: E402


OUTPUT_DIR = ROOT / "outputs" / "转债相对万得全A五分位_20260807"
MATRIX_PATH = OUTPUT_DIR / "转债相对万得全A五分位矩阵.png"
HISTORY_PATH = OUTPUT_DIR / "转债超额状态历史行情对比.png"
START = pd.Timestamp("2019-01-01")
HORIZONS = (20, 60, 120)
QUINTILE_LABELS = ["0%–20%", "20%–40%", "40%–60%", "60%–80%", "80%–100%"]


def chinese_font() -> fm.FontProperties:
    for path in [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
    ]:
        if Path(path).exists():
            return fm.FontProperties(fname=path)
    return fm.FontProperties()


FONT = chinese_font()


def load_analysis() -> pd.DataFrame:
    close, stock_close, indices = load_daily_panel()
    cb_price_median = close.median(axis=1, skipna=True)
    stock_return_120 = stock_close.pct_change(120, fill_method=None).mean(axis=1, skipna=True)
    frame = pd.concat(
        [
            historical_percentile(cb_price_median).rename("price_pct"),
            historical_percentile(stock_return_120).rename("stock_pct"),
            indices.iloc[:, 4].rename("cb_index"),
            indices.iloc[:, 0].rename("wind_all_a"),
        ],
        axis=1,
    ).dropna()
    frame = frame.loc[frame.index >= START].copy()
    frame["price_q"] = pd.cut(
        frame["price_pct"], [0, 20, 40, 60, 80, 100], labels=[1, 2, 3, 4, 5], include_lowest=True
    )
    frame["stock_q"] = pd.cut(
        frame["stock_pct"], [0, 20, 40, 60, 80, 100], labels=[1, 2, 3, 4, 5], include_lowest=True
    )
    for horizon in HORIZONS:
        cb_return = frame["cb_index"].shift(-horizon).div(frame["cb_index"]).sub(1)
        equity_return = frame["wind_all_a"].shift(-horizon).div(frame["wind_all_a"]).sub(1)
        frame[f"excess_{horizon}"] = cb_return.sub(equity_return).mul(100)
    return frame


def matrix_stat(frame: pd.DataFrame, horizon: int, stat: str) -> pd.DataFrame:
    values = frame.pivot_table(
        index="price_q",
        columns="stock_q",
        values=f"excess_{horizon}",
        aggfunc=stat,
        observed=True,
    )
    return values.reindex(index=[1, 2, 3, 4, 5], columns=[1, 2, 3, 4, 5])


def configure_style() -> None:
    plt.rcParams.update(
        {
            "axes.unicode_minus": False,
            "font.family": FONT.get_name(),
            "font.size": 11,
            "axes.labelcolor": "#28313c",
            "xtick.color": "#4b5563",
            "ytick.color": "#4b5563",
        }
    )


def plot_matrices(frame: pd.DataFrame) -> None:
    configure_style()
    fig, axes = plt.subplots(1, 3, figsize=(17, 7.2), dpi=180, constrained_layout=False)
    fig.patch.set_facecolor("white")
    fig.suptitle(
        "转债相对万得全A超额收益：五分位矩阵",
        x=0.055,
        y=0.965,
        ha="left",
        fontsize=18,
        fontproperties=FONT,
    )
    fig.text(
        0.055,
        0.91,
        "纵轴为转债价格中位数截至当日历史分位，横轴为正股120日收益截至当日历史分位；统计期为2019年至今",
        color="#687383",
        fontsize=10.5,
        fontproperties=FONT,
    )

    image = None
    for axis, horizon in zip(axes, HORIZONS):
        mean = matrix_stat(frame, horizon, "mean")
        count = matrix_stat(frame, horizon, "count")
        win = frame.assign(win=frame[f"excess_{horizon}"].gt(0)).pivot_table(
            index="price_q", columns="stock_q", values="win", aggfunc="mean", observed=True
        ).reindex(index=[1, 2, 3, 4, 5], columns=[1, 2, 3, 4, 5]).mul(100)

        display_mean = mean.iloc[::-1]
        display_count = count.iloc[::-1]
        display_win = win.iloc[::-1]
        image = axis.imshow(display_mean.to_numpy(dtype=float), cmap="RdBu_r", vmin=-6, vmax=6, aspect="equal")

        axis.set_title(f"未来{horizon}个交易日", pad=10, fontsize=13, fontproperties=FONT)
        axis.set_xticks(range(5), QUINTILE_LABELS, rotation=28, ha="right", fontproperties=FONT)
        axis.set_yticks(range(5), QUINTILE_LABELS[::-1], fontproperties=FONT)
        axis.set_xlabel("正股120日收益分位", fontproperties=FONT)
        if horizon == HORIZONS[0]:
            axis.set_ylabel("转债价格中位数分位", fontproperties=FONT)

        for row in range(5):
            for col in range(5):
                n_value = display_count.iloc[row, col]
                mean_value = display_mean.iloc[row, col]
                win_value = display_win.iloc[row, col]
                if pd.isna(n_value) or n_value < 20:
                    axis.add_patch(
                        patches.Rectangle(
                            (col - 0.5, row - 0.5), 1, 1, facecolor="#eceff3", edgecolor="white", linewidth=1.2
                        )
                    )
                    text = f"样本不足\nn={0 if pd.isna(n_value) else int(n_value)}"
                    color = "#7b8490"
                else:
                    text = f"{mean_value:+.2f}%\n胜率{win_value:.0f}%\nn={int(n_value)}"
                    color = "white" if abs(mean_value) >= 3.2 else "#202833"
                axis.text(col, row, text, ha="center", va="center", fontsize=8.3, color=color, fontproperties=FONT)

        # Stable cross-horizon cell: price Q3 (40%-60%), stock Q2 (20%-40%).
        axis.add_patch(
            patches.Rectangle((0.5, 1.5), 1, 1, fill=False, edgecolor="#f3b61f", linewidth=3.0)
        )
        axis.set_xticks(np.arange(-0.5, 5, 1), minor=True)
        axis.set_yticks(np.arange(-0.5, 5, 1), minor=True)
        axis.grid(which="minor", color="white", linewidth=1.2)
        axis.tick_params(which="minor", bottom=False, left=False)
        axis.spines[:].set_visible(False)

    colorbar_axis = fig.add_axes([0.92, 0.20, 0.012, 0.56])
    colorbar = fig.colorbar(image, cax=colorbar_axis)
    colorbar.set_label("平均超额收益（%）", fontproperties=FONT)
    for label in colorbar.ax.get_yticklabels():
        label.set_fontproperties(FONT)

    fig.text(
        0.055,
        0.025,
        "黄色边框为跨期限表现最稳定的组合：转债价格40%–60%分位 × 正股120日收益20%–40%分位。灰色单元格样本少于20个交易日。",
        color="#687383",
        fontsize=9.5,
        fontproperties=FONT,
    )
    fig.subplots_adjust(left=0.055, right=0.90, top=0.83, bottom=0.23, wspace=0.22)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(MATRIX_PATH, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def state_windows(mask: pd.Series) -> list[tuple[pd.Timestamp, pd.Timestamp, int]]:
    groups = mask.ne(mask.shift()).cumsum()
    windows: list[tuple[pd.Timestamp, pd.Timestamp, int]] = []
    for _, group in mask.groupby(groups):
        if bool(group.iloc[0]):
            windows.append((group.index[0], group.index[-1], int(group.sum())))
    return windows


def plot_history(frame: pd.DataFrame) -> None:
    configure_style()
    signal = frame["price_pct"].gt(40) & frame["price_pct"].le(60) & frame["stock_pct"].gt(20) & frame["stock_pct"].le(40)
    windows = state_windows(signal)

    normalized = frame[["cb_index", "wind_all_a"]].div(frame[["cb_index", "wind_all_a"]].iloc[0]).mul(100)
    excess_60 = frame["excess_60"]

    fig, (axis_index, axis_excess) = plt.subplots(
        2,
        1,
        figsize=(16, 8.7),
        dpi=180,
        sharex=True,
        gridspec_kw={"height_ratios": [1.12, 0.88], "hspace": 0.12},
    )
    fig.patch.set_facecolor("white")
    fig.suptitle(
        "转债相对万得全A超额状态：历史行情对比",
        x=0.08,
        y=0.975,
        ha="left",
        fontsize=18,
        fontproperties=FONT,
    )
    fig.text(
        0.08,
        0.935,
        "阴影：转债价格中位数40%–60%分位，且正股120日收益20%–40%分位（均按截至当日历史计算）",
        color="#687383",
        fontsize=10.5,
        fontproperties=FONT,
    )

    for start, end, _ in windows:
        for axis in (axis_index, axis_excess):
            axis.axvspan(start, end + pd.Timedelta(days=1), color="#f4c76b", alpha=0.30, lw=0, zorder=0)

    axis_index.plot(normalized.index, normalized["cb_index"], color="#1f5a99", lw=1.8, label="转债指数")
    axis_index.plot(normalized.index, normalized["wind_all_a"], color="#c65432", lw=1.8, label="万得全A")
    axis_index.set_ylabel("净值（2019-01-02=100）", fontproperties=FONT)
    axis_index.legend(loc="upper left", frameon=False, prop=FONT, ncol=2)

    excess_values = excess_60.to_numpy(dtype=float, na_value=np.nan)
    finite = np.isfinite(excess_values)
    axis_excess.plot(excess_60.index, excess_values, color="#45576a", lw=1.35, label="未来60日超额收益")
    axis_excess.axhline(0, color="#7a8490", lw=1.0)
    axis_excess.fill_between(
        excess_60.index, 0, excess_values, where=finite & (excess_values >= 0), color="#cf513d", alpha=0.22
    )
    axis_excess.fill_between(
        excess_60.index, 0, excess_values, where=finite & (excess_values < 0), color="#3a78ad", alpha=0.18
    )
    axis_excess.set_ylabel("未来60日超额（%）", fontproperties=FONT)
    axis_excess.legend(loc="upper left", frameon=False, prop=FONT)

    representative = [
        (pd.Timestamp("2022-03-08"), pd.Timestamp("2022-04-20"), 24),
        (pd.Timestamp("2022-09-15"), pd.Timestamp("2022-10-11"), 11),
        (pd.Timestamp("2023-07-20"), pd.Timestamp("2023-10-17"), 50),
    ]
    y_top = axis_excess.get_ylim()[1]
    for index, (start, end, days) in enumerate(representative):
        midpoint = start + (end - start) / 2
        label = f"{start:%Y.%m}–{end:%Y.%m}\n累计信号{days}日"
        axis_excess.annotate(
            label,
            xy=(midpoint, 0),
            xytext=(midpoint, y_top * (0.86 if index % 2 == 0 else 0.62)),
            ha="center",
            va="top",
            fontsize=9.5,
            color="#4e5967",
            fontproperties=FONT,
            arrowprops={"arrowstyle": "-", "color": "#8b96a3", "lw": 0.8},
        )

    for axis in (axis_index, axis_excess):
        axis.grid(axis="y", color="#d9dee5", lw=0.7, alpha=0.85)
        axis.spines[["top", "right"]].set_visible(False)
        axis.spines[["left", "bottom"]].set_color("#aab4c0")
        axis.margins(x=0)
        for label in axis.get_yticklabels():
            label.set_fontproperties(FONT)

    axis_excess.xaxis.set_major_locator(mdates.YearLocator())
    axis_excess.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    axis_excess.set_xlim(START - pd.Timedelta(days=16), frame.index[-1] + pd.Timedelta(days=16))
    for label in axis_excess.get_xticklabels():
        label.set_fontproperties(FONT)

    fig.text(
        0.08,
        0.02,
        "数据来源：转债个券历史序列。超额收益=转债指数未来60日收益−万得全A未来60日收益；末端60个交易日因尚无完整未来区间而留空。",
        color="#687383",
        fontsize=9.3,
        fontproperties=FONT,
    )
    fig.subplots_adjust(left=0.08, right=0.985, top=0.89, bottom=0.085)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(HISTORY_PATH, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    frame = load_analysis()
    plot_matrices(frame)
    plot_history(frame)
    print(MATRIX_PATH)
    print(HISTORY_PATH)


if __name__ == "__main__":
    main()
