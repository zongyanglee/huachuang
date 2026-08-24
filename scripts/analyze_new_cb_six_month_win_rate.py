from __future__ import annotations

from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import numpy as np
import pandas as pd


SHEET_CLOSE = "收盘价"
COL_CODE = "__row_id"
COL_SHEET = "__sheet_name"
COL_NAME = "转债名称"
COL_LIST = "上市日期"


def find_history_root(workspace: Path) -> Path:
    candidates = [p.parent.parent for p in workspace.rglob("*.parquet") if p.parent.name == "_special"]
    if len(candidates) != 1:
        raise RuntimeError(f"无法唯一定位历史序列目录: {candidates}")
    return candidates[0]


def load_metadata(history_root: Path) -> pd.DataFrame:
    path = next((history_root / "_special").glob("*.parquet"))
    meta = pd.read_parquet(path)
    keep = [COL_CODE, COL_NAME, COL_LIST]
    meta = meta.loc[:, keep].copy()
    meta[COL_LIST] = pd.to_datetime(meta[COL_LIST], errors="coerce")
    meta = meta.dropna(subset=[COL_CODE, COL_LIST]).drop_duplicates(COL_CODE)
    return meta.set_index(COL_CODE)


def load_close_prices(history_root: Path) -> pd.DataFrame:
    monthly = sorted(
        p
        for p in history_root.glob("[0-9][0-9][0-9][0-9]/*.parquet")
        if p.stem.isdigit() and len(p.stem) == 6
    )
    if not monthly:
        raise RuntimeError("未找到月度历史行情文件")

    blocks: list[pd.DataFrame] = []
    for path in monthly:
        block = pd.read_parquet(path, filters=[(COL_SHEET, "==", SHEET_CLOSE)])
        block = block.loc[block[COL_SHEET].eq(SHEET_CLOSE)].drop(columns=[COL_SHEET])
        block = block.set_index(COL_CODE)
        blocks.append(block)

    close = pd.concat(blocks, axis=1)
    close.columns = pd.to_datetime(close.columns)
    close = close.loc[:, ~close.columns.duplicated()].sort_index(axis=1)
    return close.apply(pd.to_numeric, errors="coerce")


def calculate_bond_results(meta: pd.DataFrame, close: pd.DataFrame) -> tuple[pd.DataFrame, pd.Timestamp]:
    market_end = close.columns.max()
    records: list[dict] = []

    for code, row in meta.iterrows():
        list_date = row[COL_LIST].normalize()
        target_date = list_date + pd.DateOffset(months=6)
        if target_date > market_end or code not in close.index:
            continue

        series = close.loc[code].dropna()
        series = series[series.gt(0)]
        if list_date not in series.index:
            continue

        exit_candidates = series[(series.index > list_date) & (series.index <= target_date)]
        if exit_candidates.empty:
            continue

        entry_price = float(series.loc[list_date])
        exit_date = exit_candidates.index[-1]
        exit_price = float(exit_candidates.iloc[-1])
        holding_return = exit_price / entry_price - 1.0
        records.append(
            {
                "转债代码": code,
                "转债名称": row[COL_NAME],
                "上市日期": list_date,
                "上市首日收盘价": entry_price,
                "六个月目标日": target_date,
                "实际退出日": exit_date,
                "退出收盘价": exit_price,
                "持有天数": (exit_date - list_date).days,
                "六个月收益率": holding_return,
                "正收益": holding_return > 0,
                "提前退出": exit_date < target_date - pd.Timedelta(days=10),
            }
        )

    result = pd.DataFrame(records).sort_values(["上市日期", "转债代码"]).reset_index(drop=True)
    return result, market_end


def summarize_by_year(detail: pd.DataFrame, market_end: pd.Timestamp) -> pd.DataFrame:
    detail = detail.copy()
    detail["上市年份"] = detail["上市日期"].dt.year
    annual = (
        detail.groupby("上市年份", as_index=False)
        .agg(
            样本数=("正收益", "size"),
            正收益只数=("正收益", "sum"),
            正收益概率=("正收益", "mean"),
            收益率中位数=("六个月收益率", "median"),
            收益率均值=("六个月收益率", "mean"),
        )
        .sort_values("上市年份")
    )
    last_complete_year = (market_end - pd.DateOffset(months=6)).year - (
        (market_end - pd.DateOffset(months=6)).month < 12
    )
    annual["完整年度"] = annual["上市年份"].le(last_complete_year)
    return annual


def plot_annual(annual: pd.DataFrame, detail: pd.DataFrame, market_end: pd.Timestamp, output: Path) -> None:
    font_path = Path.cwd() / "assets/fonts/KaiTi_GB2312.ttf"
    font = FontProperties(fname=font_path) if font_path.exists() else FontProperties()
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(13.6, 7.4), dpi=180)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    x = annual["上市年份"].to_numpy()
    y = annual["正收益概率"].to_numpy() * 100
    complete = annual["完整年度"].to_numpy()
    colors = np.where(complete, "#2F6BFF", "#A9BCEB")
    bars = ax.bar(x, y, width=0.68, color=colors, edgecolor="none", zorder=3)

    ax.axhline(50, color="#8A94A6", lw=1.2, ls="--", zorder=2)
    ax.text(x[0] - 0.48, 51.5, "50%", color="#667085", fontproperties=font, fontsize=10)

    for bar, prob, n, pos, is_complete in zip(
        bars,
        annual["正收益概率"],
        annual["样本数"],
        annual["正收益只数"],
        annual["完整年度"],
    ):
        suffix = "*" if not is_complete else ""
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 1.4,
            f"{prob:.1%}{suffix}",
            ha="center",
            va="bottom",
            color="#1D2939",
            fontproperties=font,
            fontsize=11,
        )
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            max(2.2, bar.get_height() * 0.09),
            f"{int(pos)}/{int(n)}只",
            ha="center",
            va="bottom",
            color="white" if is_complete else "#344054",
            fontproperties=font,
            fontsize=9,
        )

    complete_detail = detail[detail["上市日期"].dt.year.isin(annual.loc[annual["完整年度"], "上市年份"])]
    overall = complete_detail["正收益"].mean()
    n_all = len(complete_detail)
    n_pos = int(complete_detail["正收益"].sum())

    ax.set_title(
        "新券上市首日买入并持有6个月：正收益概率",
        loc="left",
        pad=24,
        color="#101828",
        fontproperties=font,
        fontsize=20,
    )
    ax.text(
        0,
        1.02,
        f"按上市年份分组｜完整年度合计 {overall:.1%}（{n_pos}/{n_all}只）",
        transform=ax.transAxes,
        color="#475467",
        fontproperties=font,
        fontsize=11,
    )
    ax.set_ylabel("正收益概率（%）", color="#475467", fontproperties=font, fontsize=11)
    ax.set_xticks(x)
    ax.set_xticklabels([str(v) for v in x], fontproperties=font, fontsize=10)
    ax.set_ylim(0, max(100, np.ceil((y.max() + 8) / 10) * 10))
    ax.yaxis.set_major_formatter(lambda value, _: f"{value:.0f}%")
    ax.grid(axis="y", color="#E4E7EC", lw=0.8, zorder=0)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color("#D0D5DD")
    ax.tick_params(axis="y", length=0, colors="#667085")
    ax.tick_params(axis="x", length=0, colors="#475467", pad=8)

    note = (
        f"口径：上市首日收盘价买入；上市满6个月时，取目标日及之前最近交易日收盘价卖出。"
        f"提前退市/赎回券取最后可交易收盘价；不计利息、税费。行情截至{market_end:%Y-%m-%d}。"
    )
    if (~annual["完整年度"]).any():
        note += " * 为不完整年度样本。"
    fig.text(0.075, 0.02, note, color="#667085", fontproperties=font, fontsize=9)

    fig.subplots_adjust(left=0.075, right=0.985, top=0.83, bottom=0.14)
    output.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    workspace = Path.cwd()
    output_dir = workspace / "outputs" / "新券上市后六个月正收益概率"
    output_dir.mkdir(parents=True, exist_ok=True)

    history_root = find_history_root(workspace)
    meta = load_metadata(history_root)
    close = load_close_prices(history_root)
    detail, market_end = calculate_bond_results(meta, close)
    annual = summarize_by_year(detail, market_end)

    detail.to_csv(output_dir / "新券上市后六个月收益明细.csv", index=False, encoding="utf-8-sig")
    annual.to_csv(output_dir / "新券上市后六个月正收益概率_年度.csv", index=False, encoding="utf-8-sig")
    plot_annual(annual, detail, market_end, output_dir / "新券上市后六个月正收益概率.png")

    print(f"market_end={market_end:%Y-%m-%d}")
    print(f"detail_n={len(detail)}")
    print(annual.to_string(index=False))
    print(f"output_dir={output_dir}")


if __name__ == "__main__":
    main()
