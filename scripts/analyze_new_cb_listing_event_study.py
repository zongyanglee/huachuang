from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import numpy as np
import pandas as pd


COL_SHEET = "__sheet_name"
COL_CODE = "__row_id"
SHEET_PRICE = "收盘价"
SHEET_PREMIUM = "转股溢价率"
COL_NAME = "转债名称"
COL_LIST = "上市日期"
HORIZON = 120
COMPRESSION_SEARCH_START = 5
COMPRESSION_SEARCH_END = 40
COMPRESSION_TROUGH_WINDOW = 25
SMOOTH_WINDOW = 5
OBVIOUS_RISE = 0.05
SUSTAIN_DAYS = 3


@dataclass(frozen=True)
class Cohort:
    key: str
    label: str
    start: pd.Timestamp


COHORTS = (
    Cohort("since_2017", "2017年以来", pd.Timestamp("2017-01-01")),
    Cohort("since_2024_10", "2024年10月以来", pd.Timestamp("2024-10-01")),
)


def find_history_root(workspace: Path) -> Path:
    candidates = [p.parent.parent for p in workspace.rglob("*.parquet") if p.parent.name == "_special"]
    if len(candidates) != 1:
        raise RuntimeError(f"无法唯一定位历史序列目录: {candidates}")
    return candidates[0]


def load_metadata(history_root: Path) -> pd.DataFrame:
    path = next((history_root / "_special").glob("*.parquet"))
    meta = pd.read_parquet(path)[[COL_CODE, COL_NAME, COL_LIST]].copy()
    meta[COL_LIST] = pd.to_datetime(meta[COL_LIST], errors="coerce")
    return meta.dropna(subset=[COL_CODE, COL_LIST]).drop_duplicates(COL_CODE).set_index(COL_CODE)


def load_two_sheets(history_root: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    monthly = sorted(
        p
        for p in history_root.glob("[0-9][0-9][0-9][0-9]/*.parquet")
        if p.stem.isdigit() and len(p.stem) == 6
    )
    blocks = {SHEET_PRICE: [], SHEET_PREMIUM: []}
    filters = [[(COL_SHEET, "==", SHEET_PRICE)], [(COL_SHEET, "==", SHEET_PREMIUM)]]
    for path in monthly:
        raw = pd.read_parquet(path, filters=filters)
        for sheet in blocks:
            part = raw.loc[raw[COL_SHEET].eq(sheet)].drop(columns=[COL_SHEET]).set_index(COL_CODE)
            blocks[sheet].append(part)

    result = {}
    for sheet, parts in blocks.items():
        wide = pd.concat(parts, axis=1)
        wide.columns = pd.to_datetime(wide.columns)
        wide = wide.loc[:, ~wide.columns.duplicated()].sort_index(axis=1)
        result[sheet] = wide.apply(pd.to_numeric, errors="coerce")
    return result[SHEET_PRICE], result[SHEET_PREMIUM]


def build_event_panel(
    meta: pd.DataFrame,
    price: pd.DataFrame,
    premium: pd.DataFrame,
    cohort: Cohort,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    records: list[dict] = []
    included: list[dict] = []
    codes = meta.index[meta[COL_LIST].ge(cohort.start)]

    for code in codes:
        if code not in price.index or code not in premium.index:
            continue
        list_date = meta.at[code, COL_LIST].normalize()
        p = price.loc[code]
        prem = premium.loc[code]
        valid_dates = p.index[(p.notna()) & p.gt(0) & p.index.to_series(index=p.index).ge(list_date)]
        if list_date not in valid_dates or len(valid_dates) < HORIZON + 1:
            continue
        dates = valid_dates[: HORIZON + 1]
        p0 = float(p.loc[list_date])
        prem0 = prem.loc[list_date]
        if pd.isna(prem0):
            continue
        prem0 = float(prem0)
        included.append(
            {
                "样本": cohort.label,
                "转债代码": code,
                "转债名称": meta.at[code, COL_NAME],
                "上市日期": list_date,
                "上市首日收盘价": p0,
                "上市首日转股溢价率_pct": prem0,
            }
        )
        for day, date in enumerate(dates):
            pv = float(p.loc[date])
            premv = prem.loc[date]
            records.append(
                {
                    "样本": cohort.label,
                    "转债代码": code,
                    "上市日期": list_date,
                    "交易日": day,
                    "日期": date,
                    "价格较上市日涨跌幅": pv / p0 - 1.0,
                    "转股溢价率_pct": float(premv) if pd.notna(premv) else np.nan,
                    "溢价率较上市日变动_pct": float(premv) - prem0 if pd.notna(premv) else np.nan,
                    "溢价率较上市日相对变动": float(premv) / prem0 - 1.0
                    if pd.notna(premv) and abs(prem0) > 1e-9
                    else np.nan,
                }
            )

    panel = pd.DataFrame(records)
    sample = pd.DataFrame(included)
    return panel, sample


def aggregate_path(panel: pd.DataFrame) -> pd.DataFrame:
    path = (
        panel.groupby("交易日", as_index=False)
        .agg(
            样本数=("转债代码", "nunique"),
            价格涨跌幅中位数=("价格较上市日涨跌幅", "median"),
            价格涨跌幅均值=("价格较上市日涨跌幅", "mean"),
            溢价率变动中位数_pct=("溢价率较上市日变动_pct", "median"),
            溢价率变动均值_pct=("溢价率较上市日变动_pct", "mean"),
            溢价率相对变动中位数=("溢价率较上市日相对变动", "median"),
        )
        .sort_values("交易日")
    )
    path["溢价率变动中位数_平滑_pct"] = (
        path["溢价率变动中位数_pct"].rolling(SMOOTH_WINDOW, center=True, min_periods=3).mean()
    )
    path["价格涨跌幅中位数_平滑"] = (
        path["价格涨跌幅中位数"].rolling(SMOOTH_WINDOW, center=True, min_periods=3).mean()
    )
    return path


def locate_events(path: pd.DataFrame, panel: pd.DataFrame, cohort: Cohort, sample_n: int) -> dict:
    search = path[path["交易日"].between(COMPRESSION_SEARCH_START, COMPRESSION_SEARCH_END)].dropna(
        subset=["溢价率变动中位数_平滑_pct"]
    )
    compression_start_day = int(search.loc[search["溢价率变动中位数_平滑_pct"].idxmax(), "交易日"])
    trough_search = path[
        path["交易日"].between(compression_start_day + 1, compression_start_day + COMPRESSION_TROUGH_WINDOW)
    ].dropna(subset=["溢价率变动中位数_平滑_pct"])
    compression_day = int(
        trough_search.loc[trough_search["溢价率变动中位数_平滑_pct"].idxmin(), "交易日"]
    )
    at_start = path.loc[path["交易日"].eq(compression_start_day)].iloc[0]
    at_compression = path.loc[path["交易日"].eq(compression_day)].iloc[0]
    base_price_return = float(at_compression["价格涨跌幅中位数_平滑"])

    after = path[path["交易日"].gt(compression_day)].copy()
    after["较压缩点价格涨幅"] = (1 + after["价格涨跌幅中位数_平滑"]) / (1 + base_price_return) - 1
    qualifies = after["较压缩点价格涨幅"].ge(OBVIOUS_RISE)
    sustained = qualifies.rolling(SUSTAIN_DAYS, min_periods=SUSTAIN_DAYS).sum().eq(SUSTAIN_DAYS)
    if sustained.any():
        rise_day = int(after.loc[sustained].iloc[0]["交易日"]) - SUSTAIN_DAYS + 1
        rise_row = after.loc[after["交易日"].eq(rise_day)].iloc[0]
        rise_reached = True
    else:
        rise_row = after.loc[after["较压缩点价格涨幅"].idxmax()]
        rise_day = int(rise_row["交易日"])
        rise_reached = False

    day_panel = panel.loc[panel["交易日"].eq(compression_day)].copy()
    premium_positive_share = day_panel["溢价率较上市日变动_pct"].lt(0).mean()
    price_positive_share = panel.loc[panel["交易日"].eq(rise_day), "价格较上市日涨跌幅"].gt(0).mean()

    return {
        "样本": cohort.label,
        "样本起始日": cohort.start,
        "样本数": sample_n,
        "观察窗口_交易日": HORIZON,
        "开始压缩时点_上市后交易日": compression_start_day,
        "开始压缩时溢价率较上市日变动_pct": float(at_start["溢价率变动中位数_平滑_pct"]),
        "压缩低点_上市后交易日": compression_day,
        "从开始压缩到低点_交易日": compression_day - compression_start_day,
        "压缩低点溢价率变动中位数_pct": float(at_compression["溢价率变动中位数_pct"]),
        "压缩低点溢价率平滑变动_pct": float(at_compression["溢价率变动中位数_平滑_pct"]),
        "从局部高点压缩幅度_pct": float(
            at_compression["溢价率变动中位数_平滑_pct"] - at_start["溢价率变动中位数_平滑_pct"]
        ),
        "压缩点溢价率相对变动中位数": float(at_compression["溢价率相对变动中位数"]),
        "压缩点价格较上市日涨跌幅中位数": float(at_compression["价格涨跌幅中位数"]),
        "压缩点溢价率下降样本占比": float(premium_positive_share),
        "明显涨幅阈值": OBVIOUS_RISE,
        "明显涨幅是否达到": rise_reached,
        "明显涨幅时点_上市后交易日": rise_day,
        "压缩低点后再经过交易日": rise_day - compression_day,
        "明显涨幅时较压缩点价格涨幅": float(rise_row["较压缩点价格涨幅"]),
        "明显涨幅时价格较上市日涨跌幅中位数": float(rise_row["价格涨跌幅中位数_平滑"]),
        "明显涨幅时价格高于上市日样本占比": float(price_positive_share),
    }


def plot_cohort(path: pd.DataFrame, summary: dict, output: Path, font: FontProperties) -> None:
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12.8, 8.5), dpi=180, sharex=True)
    fig.patch.set_facecolor("white")
    for ax in (ax1, ax2):
        ax.set_facecolor("white")
        ax.grid(axis="y", color="#E4E7EC", linewidth=0.8)
        ax.spines[["top", "right", "left"]].set_visible(False)
        ax.spines["bottom"].set_color("#D0D5DD")
        ax.tick_params(axis="both", length=0, colors="#667085")

    x = path["交易日"]
    prem = path["溢价率变动中位数_平滑_pct"]
    price = path["价格涨跌幅中位数_平滑"] * 100
    start_day = summary["开始压缩时点_上市后交易日"]
    n_day = summary["压缩低点_上市后交易日"]
    rise_day = summary["明显涨幅时点_上市后交易日"]

    ax1.plot(x, prem, color="#7C3AED", linewidth=2.3)
    ax1.axhline(0, color="#98A2B3", linewidth=1.1, linestyle="--")
    ax1.axvspan(start_day, n_day, color="#7C3AED", alpha=0.08)
    ax1.axvline(start_day, color="#7C3AED", linewidth=1.1, linestyle=":")
    ax1.axvline(n_day, color="#7C3AED", linewidth=1.2, linestyle="--")
    prem_start = float(path.loc[path["交易日"].eq(start_day), "溢价率变动中位数_平滑_pct"].iloc[0])
    prem_n = float(path.loc[path["交易日"].eq(n_day), "溢价率变动中位数_平滑_pct"].iloc[0])
    ax1.scatter([start_day], [prem_start], s=42, facecolor="white", edgecolor="#7C3AED", linewidth=1.5, zorder=4)
    ax1.scatter([n_day], [prem_n], s=50, color="#7C3AED", zorder=4)
    ax1.annotate(
        f"第{start_day}日开始压缩\n较上市日{prem_start:+.1f}pct",
        (start_day, prem_start),
        xytext=(-8, 14),
        textcoords="offset points",
        ha="right",
        color="#472080",
        fontproperties=font,
        fontsize=10,
    )
    ax1.annotate(
        f"第{n_day}日低点：{prem_n:+.1f}pct\n区间压缩{summary['从局部高点压缩幅度_pct']:+.1f}pct",
        (n_day, prem_n),
        xytext=(12, -30),
        textcoords="offset points",
        color="#472080",
        fontproperties=font,
        fontsize=10,
    )
    ax1.set_ylabel("溢价率较上市日变动（pct）", fontproperties=font, color="#475467")
    ax1.set_title("转股溢价率：相对上市日的中位数变动（5日平滑）", loc="left", fontproperties=font, fontsize=13)

    ax2.plot(x, price, color="#2F6BFF", linewidth=2.3)
    ax2.axhline(0, color="#98A2B3", linewidth=1.1, linestyle="--")
    ax2.axvspan(start_day, n_day, color="#7C3AED", alpha=0.08)
    ax2.axvline(n_day, color="#7C3AED", linewidth=1.2, linestyle="--")
    ax2.axvline(rise_day, color="#2F6BFF", linewidth=1.2, linestyle="--")
    y_rise = float(path.loc[path["交易日"].eq(rise_day), "价格涨跌幅中位数_平滑"].iloc[0] * 100)
    ax2.scatter([rise_day], [y_rise], s=50, color="#2F6BFF", zorder=4)
    label = "达到5%" if summary["明显涨幅是否达到"] else "窗口内最大"
    ax2.annotate(
        f"低点后{summary['压缩低点后再经过交易日']}日{label}",
        (rise_day, y_rise),
        xytext=(-12, 14),
        textcoords="offset points",
        ha="right",
        color="#1849A9",
        fontproperties=font,
        fontsize=11,
    )
    ax2.set_ylabel("价格较上市日涨跌幅（%）", fontproperties=font, color="#475467")
    ax2.set_xlabel("上市后第N个交易日", fontproperties=font, color="#475467")
    ax2.set_title("转债价格：相对上市日的中位数涨跌幅（5日平滑）", loc="left", fontproperties=font, fontsize=13)
    ax2.set_xlim(0, HORIZON)
    ax2.set_xticks(np.arange(0, HORIZON + 1, 10))

    fig.suptitle(
        f"新券上市后的溢价压缩与价格表现｜{summary['样本']}",
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
        f"固定观察120个交易日的样本 {summary['样本数']}只；横截面中位路径",
        fontproperties=font,
        fontsize=11,
        color="#475467",
    )
    fig.text(
        0.075,
        0.015,
        "注：先在上市后第5—40个交易日内寻找溢价率中位路径的早期高点（开始压缩），再在其后25日内寻找压缩低点；明显涨幅定义为相对压缩低点上涨5%并连续维持3日。pct为百分点。",
        fontproperties=font,
        fontsize=9,
        color="#667085",
    )
    fig.subplots_adjust(left=0.075, right=0.985, top=0.88, bottom=0.10, hspace=0.33)
    output.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    workspace = Path.cwd()
    output_dir = workspace / "outputs" / "新券上市事件研究_溢价压缩与涨幅"
    output_dir.mkdir(parents=True, exist_ok=True)
    font_path = workspace / "assets/fonts/KaiTi_GB2312.ttf"
    font = FontProperties(fname=font_path) if font_path.exists() else FontProperties()
    plt.rcParams["axes.unicode_minus"] = False

    history_root = find_history_root(workspace)
    meta = load_metadata(history_root)
    price, premium = load_two_sheets(history_root)

    summaries = []
    all_samples = []
    for cohort in COHORTS:
        panel, sample = build_event_panel(meta, price, premium, cohort)
        path = aggregate_path(panel)
        summary = locate_events(path, panel, cohort, len(sample))
        summaries.append(summary)
        all_samples.append(sample)
        path.to_csv(output_dir / f"{cohort.key}_事件路径.csv", index=False, encoding="utf-8-sig")
        panel.to_parquet(output_dir / f"{cohort.key}_个券事件面板.parquet", index=False)
        plot_cohort(path, summary, output_dir / f"{cohort.key}_溢价压缩与价格涨幅.png", font)

    pd.DataFrame(summaries).to_csv(output_dir / "结论汇总.csv", index=False, encoding="utf-8-sig")
    pd.concat(all_samples, ignore_index=True).to_csv(output_dir / "纳入样本.csv", index=False, encoding="utf-8-sig")
    print(pd.DataFrame(summaries).to_string(index=False))
    print(f"output_dir={output_dir}")


if __name__ == "__main__":
    main()
