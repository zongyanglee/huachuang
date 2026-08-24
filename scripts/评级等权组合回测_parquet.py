#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按债项评级分组的转债等权组合回测（parquet 数据源）。"""

from __future__ import annotations

from dataclasses import dataclass
from copy import copy
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


RATINGS = ["A", "A+", "AA-", "AA", "AA+", "AAA"]
START_DATE = pd.Timestamp("2025-01-01")
PARQUET_ROOT_NAME = "data/转债个券历史序列"
OUTPUT_DIR_NAME = "评级等权组合回测_202501至今"
RATING_SHEET = "债项评级"
RETURN_SHEET = "涨跌幅"
CLOSE_SHEET = "收盘价"
STATUS_SHEET = "交易状态"
VALID_STATUS = {"交易", "新股上市"}


@dataclass(frozen=True)
class BacktestResult:
    nav: pd.DataFrame
    daily_return: pd.DataFrame
    holding_count: pd.DataFrame
    daily_holdings_by_rating: dict[str, pd.DataFrame]
    weights: pd.DataFrame
    holdings: pd.DataFrame
    summary: pd.DataFrame
    latest_date: pd.Timestamp


def find_parquet_root(cwd: Path) -> Path:
    direct = cwd / PARQUET_ROOT_NAME
    if direct.exists():
        return direct
    candidates = [p.parent for p in cwd.glob("*/2025/202501.parquet")]
    if not candidates:
        raise FileNotFoundError("未找到转债个券历史序列 parquet 目录")
    return candidates[0]


def iter_month_files(root: Path, start: pd.Timestamp) -> list[Path]:
    files: list[Path] = []
    for year_dir in sorted(p for p in root.iterdir() if p.is_dir() and p.name.isdigit()):
        if int(year_dir.name) < start.year:
            continue
        files.extend(sorted(year_dir.glob("*.parquet")))
    return files


def read_metric_wide(root: Path, metric: str, start: pd.Timestamp) -> pd.DataFrame:
    parts: list[pd.DataFrame] = []
    for file in iter_month_files(root, start):
        df = pd.read_parquet(file)
        sub = df[df["__sheet_name"].eq(metric)].copy()
        if sub.empty:
            continue

        date_cols = [c for c in sub.columns if c not in {"__sheet_name", "__row_id"}]
        parsed = pd.to_datetime(pd.Index(date_cols), errors="coerce")
        keep_cols = [c for c, ts in zip(date_cols, parsed) if pd.notna(ts) and pd.Timestamp(ts) >= start]
        if not keep_cols:
            continue

        wide = sub.set_index("__row_id")[keep_cols].T
        wide.index = pd.to_datetime(wide.index)
        wide.index.name = "日期"
        wide.columns = wide.columns.astype(str)
        parts.append(wide)

    if not parts:
        raise ValueError(f"未读取到指标：{metric}")

    out = pd.concat(parts, axis=0)
    out = out[~out.index.duplicated(keep="last")].sort_index()
    return out


def clean_rating(df: pd.DataFrame) -> pd.DataFrame:
    return (
        df.astype("string")
        .replace({"0": pd.NA, "": pd.NA, "nan": pd.NA, "None": pd.NA})
        .apply(lambda s: s.str.strip().str.upper().str.replace("＋", "+", regex=False).str.replace("－", "-", regex=False))
    )


def max_drawdown(nav: pd.Series) -> float:
    running_max = nav.cummax()
    drawdown = nav / running_max - 1
    return float(drawdown.min())


def calc_summary(nav: pd.DataFrame, daily_return: pd.DataFrame, counts: pd.DataFrame) -> pd.DataFrame:
    rows = []
    trading_days = max(len(nav.index), 1)
    annual_factor = 252
    for rating in nav.columns:
        r = daily_return[rating].dropna()
        if r.empty:
            ann_ret = ann_vol = sharpe = mdd = np.nan
        else:
            total_ret = nav[rating].iloc[-1] / nav[rating].iloc[0] - 1
            ann_ret = (1 + total_ret) ** (annual_factor / trading_days) - 1
            ann_vol = r.std(ddof=0) * np.sqrt(annual_factor)
            sharpe = ann_ret / ann_vol if ann_vol and not np.isnan(ann_vol) else np.nan
            mdd = max_drawdown(nav[rating])
        rows.append(
            {
                "评级": rating,
                "期初净值": nav[rating].iloc[0],
                "最新净值": nav[rating].iloc[-1],
                "累计收益": nav[rating].iloc[-1] / nav[rating].iloc[0] - 1,
                "年化收益": ann_ret,
                "年化波动": ann_vol,
                "夏普(无风险0)": sharpe,
                "最大回撤": mdd,
                "最新持仓数量": counts[rating].iloc[-1],
                "平均持仓数量": counts[rating].mean(),
            }
        )
    return pd.DataFrame(rows)


def build_daily_holding_page(mask: pd.DataFrame, counts: pd.Series) -> pd.DataFrame:
    rows: list[list[str]] = []
    max_count = int(counts.max()) if not counts.empty else 0
    holding_cols = [f"持仓{i}" for i in range(1, max_count + 1)]

    for date, row in mask.iterrows():
        codes = row.index[row.to_numpy(dtype=bool)].astype(str).tolist()
        rows.append([date, len(codes), *codes, *[""] * (max_count - len(codes))])

    return pd.DataFrame(rows, columns=["日期", "持仓数量", *holding_cols])


def run_backtest(cwd: Path) -> BacktestResult:
    root = find_parquet_root(cwd)
    returns_pct = read_metric_wide(root, RETURN_SHEET, START_DATE).apply(pd.to_numeric, errors="coerce")
    close = read_metric_wide(root, CLOSE_SHEET, START_DATE).apply(pd.to_numeric, errors="coerce")
    ratings = clean_rating(read_metric_wide(root, RATING_SHEET, START_DATE))
    status = read_metric_wide(root, STATUS_SHEET, START_DATE).astype("string").apply(lambda s: s.str.strip())

    common_dates = returns_pct.index.intersection(close.index).intersection(ratings.index).intersection(status.index)
    common_cols = returns_pct.columns.intersection(close.columns).intersection(ratings.columns).intersection(status.columns)
    returns = returns_pct.loc[common_dates, common_cols] / 100.0
    close = close.loc[common_dates, common_cols]
    ratings = ratings.loc[common_dates, common_cols]
    status = status.loc[common_dates, common_cols]

    valid_trade = status.isin(VALID_STATUS) & close.notna() & returns.notna()
    daily_return = pd.DataFrame(index=common_dates)
    holding_count = pd.DataFrame(index=common_dates)
    daily_holdings_by_rating: dict[str, pd.DataFrame] = {}
    weight_frames: list[pd.DataFrame] = []
    holding_rows: list[pd.DataFrame] = []

    for rating in RATINGS:
        mask = ratings.eq(rating) & valid_trade
        counts = mask.sum(axis=1).astype(int)
        group_returns = returns.where(mask).mean(axis=1, skipna=True).fillna(0.0)
        daily_return[rating] = group_returns
        holding_count[rating] = counts
        daily_holdings_by_rating[rating] = build_daily_holding_page(mask, counts)

        weights = mask.astype("float64").div(counts.replace(0, np.nan), axis=0).fillna(0.0)
        weights = weights.loc[:, weights.ne(0).any(axis=0)].copy()
        weights = pd.concat([pd.Series(rating, index=weights.index, name="评级"), weights], axis=1)
        weight_frames.append(weights.reset_index())

        stacked = mask.stack()
        selected = stacked[stacked].index.to_frame(index=False)
        if not selected.empty:
            selected.columns = ["日期", "代码"]
            selected.insert(1, "评级", rating)
            selected["权重"] = selected["日期"].map(1 / counts.replace(0, np.nan))
            holding_rows.append(selected)

    nav = (1.0 + daily_return).cumprod()
    if not nav.empty:
        nav.iloc[0] = 1.0

    summary = calc_summary(nav, daily_return, holding_count)
    weights_all = pd.concat(weight_frames, ignore_index=True) if weight_frames else pd.DataFrame()
    holdings = pd.concat(holding_rows, ignore_index=True) if holding_rows else pd.DataFrame(columns=["日期", "评级", "代码", "权重"])

    return BacktestResult(
        nav=nav,
        daily_return=daily_return,
        holding_count=holding_count,
        daily_holdings_by_rating=daily_holdings_by_rating,
        weights=weights_all,
        holdings=holdings,
        summary=summary,
        latest_date=common_dates.max(),
    )


def save_outputs(result: BacktestResult, cwd: Path) -> tuple[Path, Path]:
    out_dir = cwd / OUTPUT_DIR_NAME
    out_dir.mkdir(parents=True, exist_ok=True)
    end_tag = result.latest_date.strftime("%Y%m%d")
    xlsx_path = out_dir / f"评级等权组合回测_20250101-{end_tag}.xlsx"
    png_path = out_dir / f"评级等权组合净值曲线_20250101-{end_tag}.png"

    with pd.ExcelWriter(xlsx_path, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        result.summary.to_excel(writer, sheet_name="绩效汇总", index=False)
        result.nav.to_excel(writer, sheet_name="净值曲线")
        result.daily_return.to_excel(writer, sheet_name="日收益率")
        result.holding_count.to_excel(writer, sheet_name="持仓数量")
        result.holdings.to_excel(writer, sheet_name="持仓明细", index=False)
        for rating in RATINGS:
            result.daily_holdings_by_rating[rating].to_excel(writer, sheet_name=f"持仓_{rating}", index=False)
        result.weights.to_excel(writer, sheet_name="等权权重", index=False)

        workbook = writer.book
        for ws in workbook.worksheets:
            ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 13
            for cell in ws[1]:
                font = copy(cell.font)
                font.bold = True
                cell.font = font

        summary_ws = workbook["绩效汇总"]
        widths = {"A": 10, "B": 12, "C": 12, "D": 13, "E": 13, "F": 13, "G": 14, "H": 13, "I": 14, "J": 14}
        for col, width in widths.items():
            summary_ws.column_dimensions[col].width = width
        for row in summary_ws.iter_rows(min_row=2, min_col=2, max_col=3):
            for cell in row:
                cell.number_format = "0.0000"
        for row in summary_ws.iter_rows(min_row=2, min_col=4, max_col=8):
            for cell in row:
                cell.number_format = "0.00%"
        for row in summary_ws.iter_rows(min_row=2, min_col=9, max_col=10):
            for cell in row:
                cell.number_format = "0.00"

        for sheet_name in ["净值曲线", "日收益率", "持仓数量"]:
            ws = workbook[sheet_name]
            for col_idx in range(2, len(RATINGS) + 2):
                col_letter = ws.cell(row=1, column=col_idx).column_letter
                ws.column_dimensions[col_letter].width = 12
                fmt = "0.00%" if sheet_name == "日收益率" else "0.0000"
                for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
                    for item in cell:
                        item.number_format = fmt

        holdings_ws = workbook["持仓明细"]
        holdings_ws.column_dimensions["A"].width = 13
        holdings_ws.column_dimensions["B"].width = 10
        holdings_ws.column_dimensions["C"].width = 12
        holdings_ws.column_dimensions["D"].width = 12
        for cell in holdings_ws.iter_cols(min_col=4, max_col=4, min_row=2, max_row=holdings_ws.max_row):
            for item in cell:
                item.number_format = "0.00%"

        for rating in RATINGS:
            ws = workbook[f"持仓_{rating}"]
            ws.column_dimensions["A"].width = 13
            ws.column_dimensions["B"].width = 10
            for col_idx in range(3, ws.max_column + 1):
                ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 12

    plt.rcParams["font.sans-serif"] = ["KaiTi", "SimHei", "Microsoft YaHei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False
    fig, ax = plt.subplots(figsize=(12, 7), dpi=160)
    for rating in RATINGS:
        ax.plot(result.nav.index, result.nav[rating], label=rating, linewidth=1.8)
    ax.set_title(f"债项评级等权组合净值曲线（2025-01-01至{result.latest_date:%Y-%m-%d}）")
    ax.set_xlabel("日期")
    ax.set_ylabel("净值")
    ax.grid(True, alpha=0.25)
    ax.legend(ncol=3, frameon=False)
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(png_path, bbox_inches="tight")
    plt.close(fig)

    return xlsx_path, png_path


def main() -> None:
    cwd = Path.cwd()
    result = run_backtest(cwd)
    xlsx_path, png_path = save_outputs(result, cwd)
    print(f"最新交易日: {result.latest_date:%Y-%m-%d}")
    print(result.summary.to_string(index=False))
    print(f"Excel: {xlsx_path}")
    print(f"PNG: {png_path}")


if __name__ == "__main__":
    main()
