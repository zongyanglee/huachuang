from __future__ import annotations

from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT / "data/转债个券历史序列"
MASTER_PATH = DATA_ROOT / "_special" / "总表.parquet"
FONT_PATH = ROOT / "assets/fonts/KaiTi_GB2312.ttf"
OUTPUT_PATH = ROOT / "outputs" / "2019年以来评级分类余额加权转股溢价率.png"
START_DATE = pd.Timestamp("2019-01-01")
RATING_ORDER = ["AAA", "AA+", "AA", "AA-", "A+", "A及以下"]
COLORS = {
    "AAA": "#1F4E79",
    "AA+": "#2E75B6",
    "AA": "#70AD47",
    "AA-": "#ED7D31",
    "A+": "#A64D79",
    "A及以下": "#C00000",
}


def read_history() -> pd.DataFrame:
    files = sorted(
        file
        for year_dir in DATA_ROOT.iterdir()
        if year_dir.is_dir() and year_dir.name.isdigit() and int(year_dir.name) >= START_DATE.year
        for file in year_dir.glob("*.parquet")
    )
    columns = ["转债代码", "交易日期", "余额", "收盘价", "转股溢价率", "债项评级"]
    parts = [pd.read_parquet(file, columns=columns) for file in files]
    if not parts:
        raise FileNotFoundError("未找到2019年以来的转债月度历史数据")
    panel = pd.concat(parts, ignore_index=True)
    panel["交易日期"] = pd.to_datetime(panel["交易日期"], errors="coerce")
    panel = panel[panel["交易日期"] >= START_DATE]
    return panel.drop_duplicates(["转债代码", "交易日期"], keep="last")


def apply_daily_cleaning(panel: pd.DataFrame) -> pd.DataFrame:
    master = pd.read_parquet(MASTER_PATH).set_index("转债代码")
    listing = pd.to_datetime(master["上市日期"], errors="coerce")
    last_trade = pd.to_datetime(master["最后交易日"], errors="coerce")

    out = panel.copy()
    out["上市日期"] = out["转债代码"].map(listing)
    out["最后交易日"] = out["转债代码"].map(last_trade)
    valid = (
        (out["上市日期"].isna() | (out["交易日期"] >= out["上市日期"]))
        & (out["最后交易日"].isna() | (out["交易日期"] <= out["最后交易日"]))
        & out["收盘价"].notna()
    )
    return out.loc[valid].copy()


def calculate_rating_weighted_premium(panel: pd.DataFrame) -> pd.DataFrame:
    out = panel.copy()
    out["余额"] = pd.to_numeric(out["余额"], errors="coerce")
    out["转股溢价率"] = pd.to_numeric(out["转股溢价率"], errors="coerce")

    rating = out["债项评级"].astype("string").str.strip()
    rating = rating.mask(rating.eq("") | rating.str.lower().eq("nan"))
    out["评级组"] = rating.where(
        rating.isin(["AAA", "AA+", "AA", "AA-", "A+"]),
        "A及以下",
    ).where(rating.notna())

    valid = out[["评级组", "转股溢价率", "余额"]].notna().all(axis=1)
    work = out.loc[valid, ["交易日期", "评级组", "转股溢价率", "余额"]].copy()
    work["加权值"] = work["转股溢价率"] * work["余额"]
    grouped = work.groupby(["交易日期", "评级组"], observed=True)[["加权值", "余额"]].sum()
    grouped["余额加权转股溢价率"] = grouped["加权值"] / grouped["余额"].replace(0, np.nan)
    result = grouped["余额加权转股溢价率"].unstack("评级组")
    return result.reindex(columns=RATING_ORDER).sort_index()


def plot_result(result: pd.DataFrame) -> None:
    fm.fontManager.addfont(str(FONT_PATH))
    font = fm.FontProperties(fname=str(FONT_PATH))
    plt.rcParams["font.family"] = font.get_name()
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(16, 8.5), dpi=180)
    for rating in RATING_ORDER:
        ax.plot(
            result.index,
            result[rating],
            color=COLORS[rating],
            linewidth=1.35,
            alpha=0.92,
            label=rating,
        )

    latest_date = result.index.max()
    ax.set_title(
        "2019年以来评级分类余额加权转股溢价率",
        loc="left",
        fontsize=22,
        fontproperties=font,
        pad=18,
    )
    ax.text(
        0,
        1.01,
        f"单位：%｜更新至 {latest_date:%Y-%m-%d}",
        transform=ax.transAxes,
        ha="left",
        va="bottom",
        fontsize=13,
        color="#595959",
        fontproperties=font,
    )
    ax.set_ylabel("余额加权转股溢价率（%）", fontsize=15, fontproperties=font)
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=(7,)))
    ax.grid(axis="y", color="#D9D9D9", linewidth=0.8, alpha=0.8)
    ax.grid(axis="x", which="major", color="#E7E6E6", linewidth=0.6, alpha=0.65)
    ax.set_xlim(result.index.min(), result.index.max())
    ax.margins(y=0.04)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#7F7F7F")
    ax.spines["bottom"].set_color("#7F7F7F")
    ax.tick_params(axis="both", labelsize=12, colors="#404040")
    legend = ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.09),
        ncol=6,
        frameon=False,
        fontsize=13,
        handlelength=2.4,
    )
    for text in legend.get_texts():
        text.set_fontproperties(font)

    fig.text(
        0.01,
        0.01,
        "注：评级分为AAA、AA+、AA、AA-、A+及A及以下；各组按当日转债余额加权。",
        ha="left",
        va="bottom",
        fontsize=11,
        color="#666666",
        fontproperties=font,
    )
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout(rect=(0, 0.055, 1, 1))
    fig.savefig(OUTPUT_PATH, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    history = read_history()
    cleaned = apply_daily_cleaning(history)
    result = calculate_rating_weighted_premium(cleaned)
    plot_result(result)
    print(f"output={OUTPUT_PATH}")
    print(f"range={result.index.min():%Y-%m-%d}..{result.index.max():%Y-%m-%d}")
    print(f"dates={len(result)}")
    print(result.tail().round(3).to_string())


if __name__ == "__main__":
    main()
