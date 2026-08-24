from __future__ import annotations

import json
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data/转债个券历史序列"
MASTER_PATH = DATA_DIR / "_special" / "总表.parquet"
FONT_PATH = ROOT / "assets/fonts/KaiTi_GB2312.ttf"
START_DATE = pd.Timestamp("2015-01-01")
BASE_LEVEL = 100.0
INDEX_COLOR = "#E6121B"


def _parse_last_trade(series: pd.Series) -> pd.Series:
    cleaned = series.astype("string").str.strip()
    cleaned = cleaned.mask(cleaned.isin(["", "0", "0.0", "None", "nan", "<NA>"]))
    return pd.to_datetime(cleaned, errors="coerce")


def _date_columns(df: pd.DataFrame) -> list[str]:
    cols: list[str] = []
    for col in df.columns:
        if col.startswith("__"):
            continue
        dt = pd.to_datetime(col, errors="coerce")
        if pd.notna(dt) and dt >= START_DATE:
            cols.append(col)
    return cols


def _load_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_parquet(path)
    sheet = raw[raw["__sheet_name"].eq(sheet_name)].copy()
    if sheet.empty:
        raise ValueError(f"{path} 未找到 sheet: {sheet_name}")
    sheet = sheet.set_index("__row_id")
    cols = _date_columns(sheet)
    return sheet[cols].apply(pd.to_numeric, errors="coerce")


def _concat_monthly(sheet_name: str) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    parquet_files = sorted(
        p
        for year_dir in DATA_DIR.iterdir()
        if year_dir.is_dir() and year_dir.name.isdigit() and int(year_dir.name) >= START_DATE.year
        for p in year_dir.glob("*.parquet")
    )
    for path in parquet_files:
        frames.append(_load_sheet(path, sheet_name))
    out = pd.concat(frames, axis=1)
    out = out.loc[:, ~out.columns.duplicated()]
    ordered_cols = sorted(out.columns, key=lambda c: pd.to_datetime(c))
    return out[ordered_cols]


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


def build_index() -> tuple[pd.DataFrame, dict[str, object]]:
    master = pd.read_parquet(MASTER_PATH).set_index("__row_id")
    listing = pd.to_datetime(master["上市日期"], errors="coerce")
    last_trade = _parse_last_trade(master["最后交易日"])

    close = _concat_monthly("收盘价")
    balance = _concat_monthly("余额")
    common_index = close.index.intersection(balance.index)
    common_cols = close.columns.intersection(balance.columns)
    close = close.loc[common_index, common_cols]
    balance = balance.loc[common_index, common_cols]

    trade_dates = pd.to_datetime(common_cols)
    close = _apply_listing_mask(close, trade_dates, listing, last_trade)
    balance = _apply_listing_mask(balance, trade_dates, listing, last_trade)

    returns = close / close.shift(1, axis=1) - 1

    records: list[dict[str, object]] = []
    level = BASE_LEVEL
    first_date = trade_dates[0]
    prev_weights = pd.Series(dtype="float64")

    for idx, (col, dt) in enumerate(zip(common_cols, trade_dates)):
        if idx == 0:
            prev_bal = pd.to_numeric(balance[col], errors="coerce")
            valid_prev = prev_bal.notna() & (prev_bal > 0) & close[col].notna()
            prev_weights = prev_bal.loc[valid_prev] / prev_bal.loc[valid_prev].sum()
            records.append({
                "date": dt.strftime("%Y-%m-%d"),
                "daily_return": 0.0,
                "index": level,
                "sample_count": int(valid_prev.sum()),
                "weight_sum": 1.0,
                "new_listing_count_excluded": 0,
            })
            continue

        ret = pd.to_numeric(returns[col], errors="coerce")
        prev_col = common_cols[idx - 1]
        prev_bal = pd.to_numeric(balance[prev_col], errors="coerce")
        prev_close = pd.to_numeric(close[prev_col], errors="coerce")
        cur_close = pd.to_numeric(close[col], errors="coerce")
        listed_today = listing.reindex(close.index).eq(dt)

        eligible = (
            prev_bal.notna()
            & (prev_bal > 0)
            & prev_close.notna()
            & cur_close.notna()
            & ret.notna()
            & (~listed_today.fillna(False))
        )
        weight_base = prev_bal.loc[eligible]
        if weight_base.sum() > 0:
            weights = weight_base / weight_base.sum()
            daily_return = float((ret.loc[eligible] * weights).sum())
            level *= 1 + daily_return
            prev_weights = weights
        else:
            daily_return = 0.0
            weights = prev_weights.iloc[0:0]

        records.append({
            "date": dt.strftime("%Y-%m-%d"),
            "daily_return": daily_return,
            "index": level,
            "sample_count": int(eligible.sum()),
            "weight_sum": float(weights.sum()) if len(weights) else 0.0,
            "new_listing_count_excluded": int(listed_today.fillna(False).sum()),
        })

    result = pd.DataFrame(records)
    audit = {
        "first_date": first_date.strftime("%Y-%m-%d"),
        "last_date": trade_dates[-1].strftime("%Y-%m-%d"),
        "rows": int(len(result)),
        "base_level": BASE_LEVEL,
        "latest_index": float(result["index"].iloc[-1]),
        "latest_daily_return": float(result["daily_return"].iloc[-1]),
    }
    return result, audit


def plot_index(result: pd.DataFrame, output_path: Path) -> None:
    if not FONT_PATH.exists():
        raise FileNotFoundError(f"未找到字体文件：{FONT_PATH}")

    font_prop = fm.FontProperties(fname=str(FONT_PATH))
    fm.fontManager.addfont(str(FONT_PATH))
    plt.rcParams["font.family"] = font_prop.get_name()
    plt.rcParams["axes.unicode_minus"] = False

    x = pd.to_datetime(result["date"])
    fig, ax = plt.subplots(figsize=(14, 7), dpi=180)
    ax.plot(x, result["index"], color=INDEX_COLOR, linewidth=2.2, label="余额加权指数")

    ax.set_ylabel("指数点位", fontproperties=font_prop, fontsize=16, color="black")
    ax.tick_params(axis="both", colors="black")
    for spine in ax.spines.values():
        spine.set_color("black")

    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=(1, 7)))
    ax.grid(True, which="major", axis="both", color="#D9D9D9", linewidth=0.8, alpha=0.75)
    ax.grid(True, which="minor", axis="x", color="#ECECEC", linewidth=0.5, alpha=0.6)
    ax.set_xlim(x.min(), x.max())

    legend = ax.legend(loc="upper left", frameon=False, prop=font_prop, fontsize=15)
    for text in legend.get_texts():
        text.set_fontproperties(font_prop)
        text.set_fontsize(15)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(font_prop)
        label.set_fontsize(14)

    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    result, audit = build_index()
    date_tag = pd.to_datetime(audit["last_date"]).strftime("%Y%m%d")
    output_dir = ROOT / "runs" / "weekly" / f"鹏华周报{date_tag}"
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / "转债余额加权指数.csv"
    png_path = output_dir / "转债余额加权指数.png"
    audit_path = output_dir / "转债余额加权指数_口径审计.json"

    result.to_csv(csv_path, index=False, encoding="utf-8-sig")
    plot_index(result, png_path)
    audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps({
        **audit,
        "output_dir": str(output_dir),
        "csv_path": str(csv_path),
        "png_path": str(png_path),
        "audit_path": str(audit_path),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
