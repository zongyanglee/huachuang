from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import pandas as pd


def main() -> None:
    workspace = Path.cwd()
    source = workspace / "outputs" / "新券上市事件日百元拟合溢价率"
    long = pd.read_csv(source / "since_2017_T1至T120_百元拟合溢价率.csv")
    recent = pd.read_csv(source / "since_2024_10_T1至T120_百元拟合溢价率.csv")
    font = FontProperties(fname=workspace / "assets/fonts/KaiTi_GB2312.ttf")
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(13.4, 6.9), dpi=200)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    ax.plot(
        long["事件日"],
        long["百元拟合溢价率_pct"],
        color="#2F6BFF",
        linewidth=2.0,
        label="2017年以来",
        zorder=3,
    )
    ax.plot(
        recent["事件日"],
        recent["百元拟合溢价率_pct"],
        color="#7C3AED",
        linewidth=2.0,
        label="2024年10月以来",
        zorder=3,
    )

    for data, color, label, offset in (
        (long, "#2F6BFF", "2017年以来", -24),
        (recent, "#7C3AED", "2024年10月以来", 12),
    ):
        for day in (1, 120):
            row = data.loc[data["事件日"].eq(day)].iloc[0]
            value = float(row["百元拟合溢价率_pct"])
            ax.scatter([day], [value], s=52, color=color, zorder=5)
            ax.annotate(
                f"{label}  T+{day}：{value:.1f}%",
                (day, value),
                xytext=(10 if day == 1 else -10, offset),
                textcoords="offset points",
                ha="left" if day == 1 else "right",
                color="#344054",
                fontproperties=font,
                fontsize=10,
            )

    ax.axvspan(5, 100, color="#98A2B3", alpha=0.06, zorder=0)
    ax.text(
        52.5,
        50.5,
        "T+5至T+100",
        ha="center",
        color="#667085",
        fontproperties=font,
        fontsize=9,
    )

    ax.set_xlim(1, 120)
    ax.set_ylim(12, 52)
    ax.set_xticks([1, 10, 20, 40, 60, 80, 100, 120])
    ax.set_xticklabels(["T+1", "T+10", "T+20", "T+40", "T+60", "T+80", "T+100", "T+120"], fontproperties=font)
    ax.set_ylabel("百元拟合溢价率（%）", fontproperties=font, color="#475467")
    ax.set_xlabel("上市后第N个交易日", fontproperties=font, color="#475467")
    ax.grid(axis="y", color="#E4E7EC", linewidth=0.8)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color("#D0D5DD")
    ax.tick_params(axis="both", length=0, colors="#667085")
    ax.legend(frameon=False, prop=font, loc="upper right")

    ax.set_title(
        "新券上市后T+1至T+120每日百元拟合溢价率",
        loc="left",
        pad=25,
        color="#101828",
        fontproperties=font,
        fontsize=20,
    )
    ax.text(
        0,
        1.02,
        "逐日原始拟合值，未经平滑",
        transform=ax.transAxes,
        color="#475467",
        fontproperties=font,
        fontsize=10,
    )
    fig.text(
        0.075,
        0.02,
        "口径：每个T+N均使用组内转债上市后第N个交易日的横截面散点独立拟合，并读取平价100元处。",
        fontproperties=font,
        fontsize=9,
        color="#667085",
    )
    fig.subplots_adjust(left=0.075, right=0.985, top=0.84, bottom=0.14)
    fig.savefig(source / "T1至T120每日百元拟合溢价率_直观对比.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    main()
