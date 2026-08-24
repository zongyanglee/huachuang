from __future__ import annotations

import json
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.ticker import PercentFormatter


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data/转债个券历史序列"
MASTER_PATH = DATA_DIR / "_special" / "总表.parquet"
FONT_PATH = ROOT / "assets/fonts/KaiTi_GB2312.ttf"
START_DATE = pd.Timestamp("2015-01-01")

GROUPS = [
    ("70以下", None, 70.0),
    ("70-90", 70.0, 90.0),
    ("90-110", 90.0, 110.0),
    ("110-130", 110.0, 130.0),
    ("130-150", 130.0, 150.0),
    ("150以上", 150.0, None),
]

COLOR_SEQUENCE = ["#E6121B", "#0262BA", "#A6A6A6", "#E6B9B8", "#B7DEE8", "#F79646"]
COLORS = {name: color for name, color in zip([group[0] for group in GROUPS], COLOR_SEQUENCE)}


def _parse_last_trade(series: pd.Series) -> pd.Series:
    cleaned = series.astype("string").str.strip()
    cleaned = cleaned.mask(cleaned.isin(["", "0", "0.0", "None", "nan", "<NA>"]))
    return pd.to_datetime(cleaned, errors="coerce")


def _load_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_parquet(path)
    part = df[df["__sheet_name"].eq(sheet_name)].copy()
    if part.empty:
        raise ValueError(f"{path} 未找到 sheet: {sheet_name}")
    part = part.set_index("__row_id")
    date_cols = []
    for col in part.columns:
        if col.startswith("__"):
            continue
        dt = pd.to_datetime(col, errors="coerce")
        if pd.notna(dt) and dt >= START_DATE:
            date_cols.append(col)
    return part[date_cols].apply(pd.to_numeric, errors="coerce")


def _apply_listing_mask(
    values: pd.DataFrame,
    trade_dates: pd.DatetimeIndex,
    listing: pd.Series,
    last_trade: pd.Series,
) -> pd.DataFrame:
    out = values.copy()
    row_listing = listing.reindex(out.index)
    row_last = last_trade.reindex(out.index)
    for col, dt in zip(out.columns, trade_dates):
        invalid = dt < row_listing
        has_last = row_last.notna()
        invalid = invalid | (has_last & (dt > row_last))
        out.loc[invalid.fillna(False), col] = pd.NA
    return out


def _percentile_rank(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.rank(method="average", pct=True)


def main() -> None:
    if not MASTER_PATH.exists():
        raise FileNotFoundError(f"未找到总表：{MASTER_PATH}")
    if not FONT_PATH.exists():
        raise FileNotFoundError(f"未找到字体文件：{FONT_PATH}")

    master = pd.read_parquet(MASTER_PATH).set_index("__row_id")
    listing = pd.to_datetime(master["上市日期"], errors="coerce")
    last_trade = _parse_last_trade(master["最后交易日"])

    parquet_files = sorted(
        p
        for year in DATA_DIR.iterdir()
        if year.is_dir() and year.name.isdigit() and int(year.name) >= START_DATE.year
        for p in year.glob("*.parquet")
    )

    rows: list[dict[str, object]] = []
    latest_date = START_DATE

    for path in parquet_files:
        parity = _load_sheet(path, "平价")
        premium = _load_sheet(path, "转股溢价率")
        balance = _load_sheet(path, "余额")

        common_index = parity.index.intersection(premium.index).intersection(balance.index)
        common_cols = parity.columns.intersection(premium.columns).intersection(balance.columns)
        parity = parity.loc[common_index, common_cols]
        premium = premium.loc[common_index, common_cols]
        balance = balance.loc[common_index, common_cols]

        trade_dates = pd.to_datetime(common_cols)
        parity = _apply_listing_mask(parity, trade_dates, listing, last_trade)
        premium = _apply_listing_mask(premium, trade_dates, listing, last_trade)
        balance = _apply_listing_mask(balance, trade_dates, listing, last_trade)

        for col, dt in zip(common_cols, trade_dates):
            if dt < START_DATE:
                continue
            latest_date = max(latest_date, dt)
            valid_balance = balance[col] > 3
            row = {"date": dt.strftime("%Y-%m-%d")}
            for name, low, high in GROUPS:
                cond = valid_balance & premium[col].notna() & parity[col].notna()
                if low is None:
                    cond = cond & (parity[col] <= high)
                elif high is None:
                    cond = cond & (parity[col] > low)
                else:
                    cond = cond & (parity[col] > low) & (parity[col] <= high)
                row[name] = premium.loc[cond, col].mean() / 100
                row[f"{name}_样本数"] = int(cond.sum())
            rows.append(row)

    avg_df = pd.DataFrame(rows).drop_duplicates("date").sort_values("date")
    pct_df = avg_df[["date"]].copy()
    for name, _, _ in GROUPS:
        avg_df[name] = pd.to_numeric(avg_df[name], errors="coerce")
        pct_df[name] = _percentile_rank(avg_df[name])
    smooth_df = pct_df[["date"]].copy()
    for name, _, _ in GROUPS:
        smooth_df[name] = pct_df[name].rolling(window=5, min_periods=1).mean()

    date_tag = latest_date.strftime("%Y%m%d")
    output_dir = ROOT / "runs" / "weekly" / f"鹏华周报{date_tag}"
    output_dir.mkdir(parents=True, exist_ok=True)
    avg_csv_path = output_dir / "平价分组平均转股溢价率_2015以来.csv"
    pct_csv_path = output_dir / "平价分组平均转股溢价率分位数.csv"
    smooth_csv_path = output_dir / "平价分组平均转股溢价率分位数_5日平滑.csv"
    png_path = output_dir / "平价分组平均转股溢价率分位数.png"

    avg_df.to_csv(avg_csv_path, index=False, encoding="utf-8-sig")
    pct_df.to_csv(pct_csv_path, index=False, encoding="utf-8-sig")
    smooth_df.to_csv(smooth_csv_path, index=False, encoding="utf-8-sig")

    font_prop = fm.FontProperties(fname=str(FONT_PATH))
    fm.fontManager.addfont(str(FONT_PATH))
    plt.rcParams["font.family"] = font_prop.get_name()
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(14, 7), dpi=180)
    x = pd.to_datetime(smooth_df["date"])
    for name, _, _ in GROUPS:
        ax.plot(x, smooth_df[name], linewidth=2.0, label=name, color=COLORS[name])

    ax.set_ylabel("历史分位数", fontproperties=font_prop, fontsize=16, color="black")
    ax.tick_params(axis="both", colors="black")
    for spine in ax.spines.values():
        spine.set_color("black")

    ax.yaxis.set_major_formatter(PercentFormatter(xmax=1, decimals=0))
    ax.set_ylim(0, 1)
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=(1, 7)))
    ax.grid(True, which="major", axis="both", color="#D9D9D9", linewidth=0.8, alpha=0.75)
    ax.grid(True, which="minor", axis="x", color="#ECECEC", linewidth=0.5, alpha=0.6)
    ax.set_xlim(x.min(), x.max())

    legend = ax.legend(loc="upper left", frameon=False, prop=font_prop, fontsize=15, ncol=2)
    for text in legend.get_texts():
        text.set_fontproperties(font_prop)
        text.set_fontsize(15)

    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(font_prop)
        label.set_fontsize(14)

    fig.tight_layout()
    fig.savefig(png_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    print(json.dumps({
        "latest_date": latest_date.strftime("%Y-%m-%d"),
        "output_dir": str(output_dir),
        "png_path": str(png_path),
        "percentile_csv_path": str(pct_csv_path),
        "smoothed_percentile_csv_path": str(smooth_csv_path),
        "average_csv_path": str(avg_csv_path),
        "rows": int(len(pct_df)),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
