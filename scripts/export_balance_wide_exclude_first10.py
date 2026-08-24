from __future__ import annotations

from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
PARQUET_ROOT = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "outputs" / "balance_wide_exclude_first10"
XLSX_PATH = OUTPUT_DIR / "余额时间序列_宽表_成交额大于零且剔除首10个交易日.xlsx"

BALANCE = "余额"
CLOSE = "收盘价"
AMOUNT = "成交额"
META_COLS = {"__sheet_name", "__row_id"}


def monthly_files() -> list[Path]:
    return sorted(
        path
        for path in PARQUET_ROOT.glob("*/*.parquet")
        if path.parent.name.isdigit() and path.stem.isdigit()
    )


def date_columns(df: pd.DataFrame) -> list[str]:
    parsed: list[tuple[pd.Timestamp, str]] = []
    for col in df.columns:
        if col in META_COLS:
            continue
        try:
            parsed.append((pd.to_datetime(col), col))
        except (TypeError, ValueError):
            continue
    return [col for _, col in sorted(parsed)]


def metric_frame(df: pd.DataFrame, metric: str, dates: list[str]) -> pd.DataFrame:
    frame = df.loc[df["__sheet_name"].eq(metric), ["__row_id", *dates]].copy()
    frame["__row_id"] = frame["__row_id"].astype(str)
    frame = frame.drop_duplicates("__row_id", keep="last").set_index("__row_id")
    return frame[dates].apply(pd.to_numeric, errors="coerce")


def transpose_metric(frame: pd.DataFrame, dates: list[str]) -> pd.DataFrame:
    result = frame.T
    result.index = pd.to_datetime(dates)
    result.index.name = "日期"
    return result


def main() -> None:
    files = monthly_files()
    if not files:
        raise SystemExit(f"未找到月度 parquet：{PARQUET_ROOT}")

    balance_parts: list[pd.DataFrame] = []
    close_parts: list[pd.DataFrame] = []
    amount_parts: list[pd.DataFrame] = []

    for path in files:
        df = pd.read_parquet(path)
        dates = date_columns(df)
        if not dates:
            continue
        available = set(df["__sheet_name"].dropna().astype(str))
        missing = {BALANCE, CLOSE, AMOUNT} - available
        if missing:
            raise KeyError(f"{path} 缺少指标：{', '.join(sorted(missing))}")

        balance_parts.append(transpose_metric(metric_frame(df, BALANCE, dates), dates))
        close_parts.append(transpose_metric(metric_frame(df, CLOSE, dates), dates))
        amount_parts.append(transpose_metric(metric_frame(df, AMOUNT, dates), dates))

    balance = pd.concat(balance_parts).sort_index()
    close = pd.concat(close_parts).sort_index()
    amount = pd.concat(amount_parts).sort_index()
    codes = sorted(set(balance.columns) | set(close.columns) | set(amount.columns))
    balance = balance.reindex(columns=codes)
    close = close.reindex(index=balance.index, columns=codes)
    amount = amount.reindex(index=balance.index, columns=codes)

    trading_day_number = close.notna().cumsum(axis=0)
    valid_mask = amount.gt(0) & close.notna() & trading_day_number.gt(10)
    balance_clean = balance.where(valid_mask)

    notes = pd.DataFrame(
        [
            ("数据来源", str(PARQUET_ROOT)),
            ("源文件数", len(files)),
            ("日期范围", f"{balance.index.min():%Y-%m-%d} 至 {balance.index.max():%Y-%m-%d}"),
            ("交易日数", len(balance.index)),
            ("转债代码数", len(codes)),
            ("宽表结构", "行=日期，列=转债代码。"),
            ("首10日定义", "每只转债自首次出现有效收盘价起，按日期排序的前10个有效交易日。"),
            ("清洗规则", "仅保留成交额大于0且交易日序号大于10的余额，其余置为空。"),
            ("余额有效单元格数", int(balance_clean.notna().sum().sum())),
        ],
        columns=["项目", "说明"],
    )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(XLSX_PATH, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        balance_clean.reset_index().to_excel(writer, sheet_name=BALANCE, index=False)
        notes.to_excel(writer, sheet_name="说明", index=False)

    print(XLSX_PATH)
    print(
        f"files={len(files)} dates={len(balance.index)} bonds={len(codes)} "
        f"balance_cells={int(balance_clean.notna().sum().sum())}"
    )


if __name__ == "__main__":
    main()
