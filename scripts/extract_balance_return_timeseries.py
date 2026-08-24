from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
PARQUET_ROOT = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "outputs" / "balance_return_timeseries"

BALANCE = "余额"
RETURN = "涨跌幅"
AMOUNT = "成交额"
META_COLS = ["__sheet_name", "__row_id"]


def iter_monthly_parquets(root: Path) -> list[Path]:
    return sorted(
        path
        for path in root.glob("*/*.parquet")
        if path.parent.name.isdigit() and path.stem.isdigit()
    )


def date_columns(df: pd.DataFrame) -> list[str]:
    cols: list[tuple[pd.Timestamp, str]] = []
    for col in df.columns:
        if col in META_COLS:
            continue
        try:
            cols.append((pd.to_datetime(col), col))
        except Exception:
            continue
    return [col for _, col in sorted(cols)]


def metric_frame(df: pd.DataFrame, metric: str, dates: list[str]) -> pd.DataFrame:
    frame = df.loc[df["__sheet_name"].eq(metric), ["__row_id", *dates]].copy()
    frame["__row_id"] = frame["__row_id"].astype(str)
    frame = frame.drop_duplicates("__row_id", keep="last").set_index("__row_id")
    return frame[dates].apply(pd.to_numeric, errors="coerce")


def to_timeseries(frame: pd.DataFrame, dates: list[str]) -> pd.DataFrame:
    out = frame.T
    out.index = pd.to_datetime(dates).strftime("%Y-%m-%d")
    out.index.name = "日期"
    return out


def main() -> None:
    files = iter_monthly_parquets(PARQUET_ROOT)
    if not files:
        raise SystemExit(f"未找到月度 parquet: {PARQUET_ROOT}")

    balance_parts: list[pd.DataFrame] = []
    return_parts: list[pd.DataFrame] = []
    positive_amount_cells = 0
    total_candidate_cells = 0

    for file in files:
        df = pd.read_parquet(file)
        dates = date_columns(df)
        if not dates:
            continue

        missing = {BALANCE, RETURN, AMOUNT} - set(df["__sheet_name"].dropna().unique())
        if missing:
            raise KeyError(f"{file} 缺少指标: {', '.join(sorted(missing))}")

        amount = metric_frame(df, AMOUNT, dates)
        balance = metric_frame(df, BALANCE, dates)
        ret = metric_frame(df, RETURN, dates)

        codes = amount.index.union(balance.index).union(ret.index)
        amount = amount.reindex(codes)
        balance = balance.reindex(codes)
        ret = ret.reindex(codes)

        traded = amount.gt(0)
        total_candidate_cells += int(traded.size)
        positive_amount_cells += int(traded.sum().sum())

        balance_parts.append(to_timeseries(balance.where(traded), dates))
        return_parts.append(to_timeseries(ret.where(traded), dates))

    balance_ts = pd.concat(balance_parts, axis=0).sort_index()
    return_ts = pd.concat(return_parts, axis=0).sort_index()

    all_codes = sorted(set(balance_ts.columns).union(return_ts.columns))
    balance_ts = balance_ts.reindex(columns=all_codes)
    return_ts = return_ts.reindex(columns=all_codes)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    balance_csv = OUTPUT_DIR / "余额_成交额大于零清洗.csv"
    return_csv = OUTPUT_DIR / "涨跌幅_成交额大于零清洗.csv"
    summary_json = OUTPUT_DIR / "summary.json"

    balance_ts.reset_index().to_csv(balance_csv, index=False, encoding="utf-8-sig")
    return_ts.reset_index().to_csv(return_csv, index=False, encoding="utf-8-sig")

    summary = {
        "source": str(PARQUET_ROOT),
        "source_file_count": len(files),
        "date_start": str(balance_ts.index.min()) if not balance_ts.empty else None,
        "date_end": str(balance_ts.index.max()) if not balance_ts.empty else None,
        "date_count": int(len(balance_ts.index)),
        "bond_count": int(len(all_codes)),
        "cleaning_rule": "仅保留同一代码、同一日期成交额>0的余额和涨跌幅；其余置为空。",
        "positive_amount_cells": positive_amount_cells,
        "total_candidate_cells": total_candidate_cells,
        "balance_non_null_cells": int(balance_ts.notna().sum().sum()),
        "return_non_null_cells": int(return_ts.notna().sum().sum()),
        "balance_csv": str(balance_csv),
        "return_csv": str(return_csv),
    }
    summary_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(summary_json)
    print(
        f"dates={summary['date_count']} bonds={summary['bond_count']} "
        f"files={summary['source_file_count']}"
    )


if __name__ == "__main__":
    main()
