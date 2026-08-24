from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import numpy as np
import pandas as pd


HORIZON = 120
SEARCH_START = 5
SEARCH_END = 40
TROUGH_WINDOW = 25
SMOOTH_WINDOW = 5
RISE_THRESHOLD = 0.05
SUSTAIN_DAYS = 3

COHORTS = (
    ("since_2017", "2017年以来"),
    ("since_2024_10", "2024年10月以来"),
)


def build_absolute_path(panel: pd.DataFrame, sample: pd.DataFrame) -> pd.DataFrame:
    initial = sample[["转债代码", "上市首日收盘价"]].drop_duplicates("转债代码")
    data = panel.merge(initial, on="转债代码", how="left", validate="many_to_one")
    data["转债价格"] = data["上市首日收盘价"] * (1 + data["价格较上市日涨跌幅"])
    path = (
        data.groupby("交易日", as_index=False)
        .agg(
            样本数=("转债代码", "nunique"),
            转股溢价率中位数_pct=("转股溢价率_pct", "median"),
            转债价格中位数=("转债价格", "median"),
        )
        .sort_values("交易日")
    )
    path["转股溢价率中位数_平滑_pct"] = (
        path["转股溢价率中位数_pct"].rolling(SMOOTH_WINDOW, center=True, min_periods=3).mean()
    )
    path["转债价格中位数_平滑"] = (
        path["转债价格中位数"].rolling(SMOOTH_WINDOW, center=True, min_periods=3).mean()
    )
    return path


def locate_events(path: pd.DataFrame, label: str, sample_n: int) -> dict:
    initial = path.loc[path["交易日"].eq(0)].iloc[0]
    search = path[path["交易日"].between(SEARCH_START, SEARCH_END)]
    start_day = int(search.loc[search["转股溢价率中位数_平滑_pct"].idxmax(), "交易日"])
    trough_search = path[path["交易日"].between(start_day + 1, start_day + TROUGH_WINDOW)]
    trough_day = int(
        trough_search.loc[trough_search["转股溢价率中位数_平滑_pct"].idxmin(), "交易日"]
    )
    start = path.loc[path["交易日"].eq(start_day)].iloc[0]
    trough = path.loc[path["交易日"].eq(trough_day)].iloc[0]
    base_price = float(trough["转债价格中位数_平滑"])

    after = path[path["交易日"].gt(trough_day)].copy()
    after["价格较压缩低点涨幅"] = after["转债价格中位数_平滑"] / base_price - 1
    qualifies = after["价格较压缩低点涨幅"].ge(RISE_THRESHOLD)
    sustained = qualifies.rolling(SUSTAIN_DAYS, min_periods=SUSTAIN_DAYS).sum().eq(SUSTAIN_DAYS)
    if sustained.any():
        rise_day = int(after.loc[sustained].iloc[0]["交易日"]) - SUSTAIN_DAYS + 1
        rise = after.loc[after["交易日"].eq(rise_day)].iloc[0]
        reached = True
    else:
        rise = after.loc[after["价格较压缩低点涨幅"].idxmax()]
        rise_day = int(rise["交易日"])
        reached = False

    return {
        "样本": label,
        "样本数": sample_n,
        "上市日转股溢价率中位数_pct": float(initial["转股溢价率中位数_pct"]),
        "上市日转债价格中位数": float(initial["转债价格中位数"]),
        "开始压缩时点_上市后交易日": start_day,
        "开始压缩时转股溢价率中位数_pct": float(start["转股溢价率中位数_平滑_pct"]),
        "开始压缩时转债价格中位数": float(start["转债价格中位数_平滑"]),
        "压缩低点_上市后交易日": trough_day,
        "压缩低点转股溢价率中位数_pct": float(trough["转股溢价率中位数_平滑_pct"]),
        "压缩低点转债价格中位数": float(trough["转债价格中位数_平滑"]),
        "局部高点至低点压缩幅度_pct": float(
            trough["转股溢价率中位数_平滑_pct"] - start["转股溢价率中位数_平滑_pct"]
        ),
        "明显涨幅是否达到": reached,
        "明显涨幅时点_上市后交易日": rise_day,
        "压缩低点后再经过交易日": rise_day - trough_day,
        "明显涨幅时转债价格中位数": float(rise["转债价格中位数_平滑"]),
        "明显涨幅时较低点价格涨幅": float(rise["价格较压缩低点涨幅"]),
    }


def plot_absolute(path: pd.DataFrame, summary: dict, output: Path, font: FontProperties) -> None:
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12.8, 8.5), dpi=180, sharex=True)
    fig.patch.set_facecolor("white")
    for ax in (ax1, ax2):
        ax.set_facecolor("white")
        ax.grid(axis="y", color="#E4E7EC", linewidth=0.8)
        ax.spines[["top", "right", "left"]].set_visible(False)
        ax.spines["bottom"].set_color("#D0D5DD")
        ax.tick_params(axis="both", length=0, colors="#667085")

    x = path["交易日"]
    premium = path["转股溢价率中位数_平滑_pct"]
    price = path["转债价格中位数_平滑"]
    start_day = summary["开始压缩时点_上市后交易日"]
    trough_day = summary["压缩低点_上市后交易日"]
    rise_day = summary["明显涨幅时点_上市后交易日"]

    ax1.plot(x, premium, color="#7C3AED", linewidth=2.3)
    ax1.axvspan(start_day, trough_day, color="#7C3AED", alpha=0.08)
    ax1.axvline(start_day, color="#7C3AED", linewidth=1.1, linestyle=":")
    ax1.axvline(trough_day, color="#7C3AED", linewidth=1.2, linestyle="--")
    ax1.scatter(
        [start_day, trough_day],
        [summary["开始压缩时转股溢价率中位数_pct"], summary["压缩低点转股溢价率中位数_pct"]],
        s=45,
        color="#7C3AED",
        zorder=4,
    )
    ax1.annotate(
        f"第{start_day}日：{summary['开始压缩时转股溢价率中位数_pct']:.1f}%",
        (start_day, summary["开始压缩时转股溢价率中位数_pct"]),
        xytext=(-10, 14),
        textcoords="offset points",
        ha="right",
        color="#472080",
        fontproperties=font,
        fontsize=10,
    )
    ax1.annotate(
        f"第{trough_day}日：{summary['压缩低点转股溢价率中位数_pct']:.1f}%",
        (trough_day, summary["压缩低点转股溢价率中位数_pct"]),
        xytext=(10, -22),
        textcoords="offset points",
        color="#472080",
        fontproperties=font,
        fontsize=10,
    )
    ax1.set_ylabel("转股溢价率中位数（%）", fontproperties=font, color="#475467")
    ax1.set_title("转股溢价率中位数（绝对值，5日平滑）", loc="left", fontproperties=font, fontsize=13)

    ax2.plot(x, price, color="#2F6BFF", linewidth=2.3)
    ax2.axvspan(start_day, trough_day, color="#7C3AED", alpha=0.08)
    ax2.axvline(trough_day, color="#7C3AED", linewidth=1.2, linestyle="--")
    ax2.axvline(rise_day, color="#2F6BFF", linewidth=1.2, linestyle="--")
    rise_price = summary["明显涨幅时转债价格中位数"]
    ax2.scatter([rise_day], [rise_price], s=50, color="#2F6BFF", zorder=4)
    label = "达到5%" if summary["明显涨幅是否达到"] else "窗口内最大"
    ax2.annotate(
        f"低点后{summary['压缩低点后再经过交易日']}日{label}\n价格中位数{rise_price:.1f}元",
        (rise_day, rise_price),
        xytext=(-10, 14),
        textcoords="offset points",
        ha="right",
        color="#1849A9",
        fontproperties=font,
        fontsize=10,
    )
    ax2.set_ylabel("转债价格中位数（元）", fontproperties=font, color="#475467")
    ax2.set_xlabel("上市后第N个交易日", fontproperties=font, color="#475467")
    ax2.set_title("转债价格中位数（绝对值，5日平滑）", loc="left", fontproperties=font, fontsize=13)
    ax2.set_xlim(0, HORIZON)
    ax2.set_xticks(np.arange(0, HORIZON + 1, 10))

    fig.suptitle(
        f"新券上市后的溢价率与价格中位数｜{summary['样本']}",
        x=0.075,
        y=0.98,
        ha="left",
        fontproperties=font,
        fontsize=20,
        color="#101828",
    )
    fig.text(
        0.075,
        0.935,
        f"固定观察120个交易日的样本 {summary['样本数']}只；均为绝对中位数，不计算相对上市日变动",
        fontproperties=font,
        fontsize=11,
        color="#475467",
    )
    fig.text(
        0.075,
        0.015,
        "注：压缩区间按绝对转股溢价率中位数的5日平滑路径识别；明显涨幅定义为价格中位数相对压缩低点上涨5%并连续维持3日。",
        fontproperties=font,
        fontsize=9,
        color="#667085",
    )
    fig.subplots_adjust(left=0.075, right=0.985, top=0.88, bottom=0.10, hspace=0.33)
    fig.savefig(output, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    workspace = Path.cwd()
    source = workspace / "outputs" / "新券上市事件研究_溢价压缩与涨幅"
    output = workspace / "outputs" / "新券上市事件研究_价格与溢价率绝对中位数"
    output.mkdir(parents=True, exist_ok=True)
    sample_all = pd.read_csv(source / "纳入样本.csv")
    font_path = workspace / "assets/fonts/KaiTi_GB2312.ttf"
    font = FontProperties(fname=font_path) if font_path.exists() else FontProperties()
    plt.rcParams["axes.unicode_minus"] = False

    summaries = []
    for key, label in COHORTS:
        panel = pd.read_parquet(source / f"{key}_个券事件面板.parquet")
        sample = sample_all.loc[sample_all["样本"].eq(label)].copy()
        path = build_absolute_path(panel, sample)
        summary = locate_events(path, label, len(sample))
        summaries.append(summary)
        path.to_csv(output / f"{key}_绝对中位数路径.csv", index=False, encoding="utf-8-sig")
        plot_absolute(path, summary, output / f"{key}_价格与溢价率绝对中位数.png", font)

    pd.DataFrame(summaries).to_csv(output / "绝对中位数结论汇总.csv", index=False, encoding="utf-8-sig")
    print(pd.DataFrame(summaries).to_string(index=False))
    print(f"output_dir={output}")


if __name__ == "__main__":
    main()
