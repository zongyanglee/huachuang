from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import numpy as np
import pandas as pd
from scipy.optimize import curve_fit


COL_SHEET = "__sheet_name"
COL_CODE = "__row_id"
COL_LIST = "上市日期"
SHEET_PARITY = "平价"
SHEET_PREMIUM = "转股溢价率"
SHEET_TURNOVER = "换手率"
HORIZON = 120
MIN_FIT_N = 8

COHORTS = (
    ("since_2017", "2017年以来", pd.Timestamp("2017-01-01")),
    ("since_2024_10", "2024年10月以来", pd.Timestamp("2024-10-01")),
)


def inverse_cubic(x: np.ndarray | float, a: float, b: float, c: float, d: float):
    return a / np.power(x, 3) + b / np.power(x, 2) + c / x + d


def find_history_root(workspace: Path) -> Path:
    candidates = [p.parent.parent for p in workspace.rglob("*.parquet") if p.parent.name == "_special"]
    if len(candidates) != 1:
        raise RuntimeError(f"无法唯一定位历史序列目录: {candidates}")
    return candidates[0]


def load_metadata(history_root: Path) -> pd.DataFrame:
    path = next((history_root / "_special").glob("*.parquet"))
    meta = pd.read_parquet(path)[[COL_CODE, COL_LIST]].copy()
    meta[COL_LIST] = pd.to_datetime(meta[COL_LIST], errors="coerce")
    return meta.dropna(subset=[COL_CODE, COL_LIST]).drop_duplicates(COL_CODE).set_index(COL_CODE)


def load_event_sheets(history_root: Path) -> dict[str, pd.DataFrame]:
    sheets = (SHEET_PARITY, SHEET_PREMIUM, SHEET_TURNOVER)
    monthly = sorted(
        p
        for p in history_root.glob("[0-9][0-9][0-9][0-9]/*.parquet")
        if p.stem.isdigit() and len(p.stem) == 6
    )
    filters = [[(COL_SHEET, "==", sheet)] for sheet in sheets]
    blocks: dict[str, list[pd.DataFrame]] = {sheet: [] for sheet in sheets}
    for path in monthly:
        raw = pd.read_parquet(path, filters=filters)
        for sheet in sheets:
            part = raw.loc[raw[COL_SHEET].eq(sheet)].drop(columns=[COL_SHEET]).set_index(COL_CODE)
            blocks[sheet].append(part)

    result = {}
    for sheet, parts in blocks.items():
        wide = pd.concat(parts, axis=1)
        wide.columns = pd.to_datetime(wide.columns)
        wide = wide.loc[:, ~wide.columns.duplicated()].sort_index(axis=1)
        result[sheet] = wide.apply(pd.to_numeric, errors="coerce")
    return result


def build_event_panel(meta: pd.DataFrame, sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    parity = sheets[SHEET_PARITY]
    premium = sheets[SHEET_PREMIUM]
    turnover = sheets[SHEET_TURNOVER]
    records: list[dict] = []
    codes = meta.index[meta[COL_LIST].ge(pd.Timestamp("2017-01-01"))]

    for code in codes:
        if code not in parity.index or code not in premium.index or code not in turnover.index:
            continue
        list_date = meta.at[code, COL_LIST].normalize()
        p = parity.loc[code]
        prem = premium.loc[code]
        turn = turnover.loc[code]
        dates = p.index[(p.notna()) & p.gt(0) & prem.notna() & p.index.to_series(index=p.index).ge(list_date)]
        if len(dates) == 0 or dates[0] != list_date:
            continue
        for event_day, date in enumerate(dates[:HORIZON], start=1):
            records.append(
                {
                    "转债代码": code,
                    "上市日期": list_date,
                    "事件日": event_day,
                    "实际日期": date,
                    "平价": float(p.loc[date]),
                    "转股溢价率": float(prem.loc[date]),
                    "换手率": float(turn.loc[date]) if pd.notna(turn.loc[date]) else np.nan,
                }
            )
    return pd.DataFrame(records)


def prepare_fit_sample(day_data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    work = day_data[["转债代码", "平价", "转股溢价率", "换手率"]].copy()
    for col in ("平价", "转股溢价率", "换手率"):
        work[col] = pd.to_numeric(work[col], errors="coerce")
    work = work.replace(0, np.nan).dropna(subset=["平价", "转股溢价率", "换手率"])
    eligible = work[(work["平价"] > 70) & (work["平价"] < 130) & (work["换手率"] < 50)].copy()
    if eligible.empty:
        return eligible, eligible
    low = eligible["转股溢价率"].quantile(0.03)
    high = eligible["转股溢价率"].quantile(0.97)
    fit = eligible[(eligible["转股溢价率"] > low) & (eligible["转股溢价率"] < high)].copy()
    return eligible, fit


def fit_event_day(day_data: pd.DataFrame, event_day: int) -> tuple[dict, pd.DataFrame, pd.DataFrame]:
    eligible, fit = prepare_fit_sample(day_data)
    row = {
        "事件日": event_day,
        "原始散点数": int(day_data["转债代码"].nunique()),
        "70至130平价且换手率低于50样本数": int(len(eligible)),
        "拟合样本数": int(len(fit)),
        "百元拟合溢价率_pct": np.nan,
        "R2": np.nan,
        "RMSE": np.nan,
        "MAE": np.nan,
        "a": np.nan,
        "b": np.nan,
        "c": np.nan,
        "d": np.nan,
    }
    if len(fit) < MIN_FIT_N:
        return row, eligible, fit
    try:
        x = fit["平价"].to_numpy(float)
        y = fit["转股溢价率"].to_numpy(float)
        popt, _ = curve_fit(inverse_cubic, x, y, maxfev=20000)
        y_hat = inverse_cubic(x, *popt)
        residual = y - y_hat
        sse = float(np.sum(residual**2))
        sst = float(np.sum((y - y.mean()) ** 2))
        row.update(
            {
                "百元拟合溢价率_pct": float(inverse_cubic(100.0, *popt)),
                "R2": float(1 - sse / sst) if sst > 0 else np.nan,
                "RMSE": float(np.sqrt(np.mean(residual**2))),
                "MAE": float(np.mean(np.abs(residual))),
                "a": float(popt[0]),
                "b": float(popt[1]),
                "c": float(popt[2]),
                "d": float(popt[3]),
            }
        )
    except Exception:
        pass
    return row, eligible, fit


def fit_cohort(panel: pd.DataFrame, start: pd.Timestamp) -> tuple[pd.DataFrame, dict[int, tuple[pd.DataFrame, pd.DataFrame]]]:
    cohort = panel.loc[panel["上市日期"].ge(start)].copy()
    rows = []
    scatter = {}
    for event_day in range(1, HORIZON + 1):
        day_data = cohort.loc[cohort["事件日"].eq(event_day)]
        row, eligible, fit = fit_event_day(day_data, event_day)
        rows.append(row)
        if event_day == 1:
            scatter[event_day] = (eligible, fit)
    result = pd.DataFrame(rows)
    result["百元拟合溢价率_5日平滑_pct"] = (
        result["百元拟合溢价率_pct"].rolling(5, center=True, min_periods=3).mean()
    )
    return result, scatter


def plot_t1_scatter(
    result_row: pd.Series,
    eligible: pd.DataFrame,
    fit: pd.DataFrame,
    label: str,
    output: Path,
    font: FontProperties,
) -> None:
    fig, ax = plt.subplots(figsize=(10.2, 6.2), dpi=180)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.scatter(eligible["平价"], eligible["转股溢价率"], s=24, alpha=0.24, color="#98A2B3", label="筛选后散点")
    ax.scatter(fit["平价"], fit["转股溢价率"], s=26, alpha=0.68, color="#2F6BFF", label="3%—97%拟合样本")
    if pd.notna(result_row["百元拟合溢价率_pct"]):
        params = [result_row[c] for c in ("a", "b", "c", "d")]
        xs = np.linspace(70, 130, 300)
        ys = inverse_cubic(xs, *params)
        ax.plot(xs, ys, color="#7C3AED", linewidth=2.2, label="反三次函数拟合")
        y100 = result_row["百元拟合溢价率_pct"]
        ax.scatter([100], [y100], s=70, color="#7C3AED", zorder=5)
        ax.annotate(
            f"平价100：{y100:.1f}%",
            (100, y100),
            xytext=(12, 12),
            textcoords="offset points",
            fontproperties=font,
            fontsize=11,
            color="#472080",
        )
    fig.suptitle(
        f"T+1上市首日散点拟合｜{label}",
        x=0.08,
        y=0.98,
        ha="left",
        fontproperties=font,
        fontsize=19,
        color="#101828",
    )
    fig.text(
        0.08,
        0.925,
        f"原始{int(result_row['原始散点数'])}只；拟合样本{int(result_row['拟合样本数'])}只；R2={result_row['R2']:.2f}",
        fontproperties=font,
        fontsize=10,
        color="#475467",
    )
    ax.set_xlabel("平价（元）", fontproperties=font, color="#475467")
    ax.set_ylabel("转股溢价率（%）", fontproperties=font, color="#475467")
    ax.grid(color="#E4E7EC", linewidth=0.8)
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.spines["bottom"].set_color("#D0D5DD")
    ax.tick_params(length=0, colors="#667085")
    ax.legend(frameon=False, prop=font, loc="upper right")
    fig.text(
        0.08,
        0.02,
        "口径：平价70—130元、换手率<50%，剔除转股溢价率两端各3%样本，反三次函数拟合并读取平价100元处。",
        fontproperties=font,
        fontsize=9,
        color="#667085",
    )
    fig.subplots_adjust(left=0.08, right=0.98, top=0.84, bottom=0.14)
    fig.savefig(output, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def plot_comparison(results: dict[str, pd.DataFrame], output: Path, font: FontProperties) -> None:
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(13.2, 8.4), dpi=180, sharex=True)
    fig.patch.set_facecolor("white")
    colors = {"2017年以来": "#2F6BFF", "2024年10月以来": "#7C3AED"}
    for label, result in results.items():
        ax1.plot(
            result["事件日"],
            result["百元拟合溢价率_pct"],
            color=colors[label],
            linewidth=0.9,
            alpha=0.24,
        )
        ax1.plot(
            result["事件日"],
            result["百元拟合溢价率_5日平滑_pct"],
            color=colors[label],
            linewidth=2.3,
            label=label,
        )
        ax2.plot(
            result["事件日"], result["拟合样本数"], color=colors[label], linewidth=2.0, label=label
        )

    for ax in (ax1, ax2):
        ax.set_facecolor("white")
        ax.grid(axis="y", color="#E4E7EC", linewidth=0.8)
        ax.spines[["top", "right", "left"]].set_visible(False)
        ax.spines["bottom"].set_color("#D0D5DD")
        ax.tick_params(axis="both", length=0, colors="#667085")

    ax1.set_title("各事件日横截面拟合得到的百元溢价率", loc="left", fontproperties=font, fontsize=13)
    ax1.set_ylabel("百元拟合溢价率（%）", fontproperties=font, color="#475467")
    ax1.legend(frameon=False, prop=font, loc="best")
    ax2.set_title("每日实际参与拟合的样本数", loc="left", fontproperties=font, fontsize=13)
    ax2.set_ylabel("拟合样本数（只）", fontproperties=font, color="#475467")
    ax2.set_xlabel("上市后事件日（T+1为上市首日）", fontproperties=font, color="#475467")
    ax2.set_xlim(1, HORIZON)
    ax2.set_xticks([1, 10, 20, 40, 60, 80, 100, 120])

    fig.suptitle(
        "新券上市后各事件日的百元拟合溢价率",
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
        "T+1按所有转债上市首日散点拟合，T+N按各券上市后第N个交易日散点拟合；细线为每日值，粗线为5日平滑。",
        fontproperties=font,
        fontsize=10,
        color="#475467",
    )
    fig.text(
        0.075,
        0.015,
        "拟合口径与现有百元拟合脚本一致：平价70—130元、换手率<50%，剔除溢价率两端各3%，反三次函数 y=a/x^3+b/x^2+c/x+d，读取x=100。",
        fontproperties=font,
        fontsize=9,
        color="#667085",
    )
    fig.subplots_adjust(left=0.075, right=0.985, top=0.88, bottom=0.10, hspace=0.33)
    fig.savefig(output, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    workspace = Path.cwd()
    output = workspace / "outputs" / "新券上市事件日百元拟合溢价率"
    output.mkdir(parents=True, exist_ok=True)
    font_path = workspace / "assets/fonts/KaiTi_GB2312.ttf"
    font = FontProperties(fname=font_path) if font_path.exists() else FontProperties()
    plt.rcParams["axes.unicode_minus"] = False

    history_root = find_history_root(workspace)
    meta = load_metadata(history_root)
    sheets = load_event_sheets(history_root)
    panel = build_event_panel(meta, sheets)
    panel.to_parquet(output / "上市事件日平价溢价率换手率面板.parquet", index=False)

    results: dict[str, pd.DataFrame] = {}
    for key, label, start in COHORTS:
        result, scatter = fit_cohort(panel, start)
        result.insert(0, "样本", label)
        result.to_csv(output / f"{key}_T1至T{HORIZON}_百元拟合溢价率.csv", index=False, encoding="utf-8-sig")
        results[label] = result
        eligible, fit = scatter[1]
        plot_t1_scatter(
            result.loc[result["事件日"].eq(1)].iloc[0],
            eligible,
            fit,
            label,
            output / f"{key}_T1上市首日散点拟合.png",
            font,
        )

    summary_days = [1, 5, 10, 20, 40, 60, 80, 100, 120]
    summary = pd.concat(
        [result.loc[result["事件日"].isin(summary_days)] for result in results.values()], ignore_index=True
    )
    summary.to_csv(output / "关键事件日汇总.csv", index=False, encoding="utf-8-sig")
    plot_comparison(results, output / "两组_T1至T120_百元拟合溢价率对比.png", font)

    print(summary[["样本", "事件日", "原始散点数", "拟合样本数", "百元拟合溢价率_pct", "R2"]].to_string(index=False))
    print(f"panel_rows={len(panel)}")
    print(f"output_dir={output}")


if __name__ == "__main__":
    main()
