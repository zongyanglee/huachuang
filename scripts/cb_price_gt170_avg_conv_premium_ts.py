from __future__ import annotations

import argparse
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator, Sequence

import numpy as np
import pandas as pd
import pyarrow.parquet as pq


@dataclass(frozen=True)
class Config:
    root: Path
    parquet_dir: Path = Path("data/转债个券历史序列")

    price_sheet: str = "收盘价"
    premium_sheet: str = "转股溢价率"
    status_sheet: str = "交易状态"

    price_threshold: float = 170.0
    trading_values: tuple[str, ...] = ("交易", "正常交易")
    ignore_status_from: str | None = None  # 此日期起不再使用交易状态过滤

    start: str | None = None  # YYYY-MM-DD
    end: str | None = None  # YYYY-MM-DD

    out: Path | None = None  # .xlsx or .csv


def _iter_monthly_parquets(root: Path, parquet_dir: Path) -> list[Path]:
    base = (root / parquet_dir).resolve()
    if not base.exists():
        raise FileNotFoundError(f"Not found: {base}")

    files: list[Path] = []
    for p in base.rglob("*.parquet"):
        if p.name == "sheet_manifest.parquet":
            continue
        # _special/总表.parquet 不参与按日计算
        if p.name == "总表.parquet":
            continue
        # 仅取 yyyy/yyyymm.parquet
        if not p.parent.name.isdigit():
            continue
        if not re.match(r"^\d{6}\.parquet$", p.name):
            continue
        files.append(p)

    files.sort(key=lambda x: str(x))
    return files


def _to_number(df: pd.DataFrame) -> pd.DataFrame:
    out = df.astype("string")
    out = out.replace({"": pd.NA})
    out = out.apply(lambda s: s.str.replace(",", "", regex=False).str.strip())
    return out.apply(pd.to_numeric, errors="coerce")


def _sheet_wide(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    sub = df.loc[df["__sheet_name"] == sheet_name].copy()
    if sub.empty:
        return pd.DataFrame()
    sub = sub.drop(columns=["__sheet_name"]).set_index("__row_id")
    sub.index = sub.index.astype("string")

    col_map: dict[object, pd.Timestamp] = {}
    for c in sub.columns:
        ts = pd.to_datetime(c, errors="coerce")
        if pd.notna(ts):
            col_map[c] = pd.Timestamp(ts)
    date_cols = [c for c in sub.columns if c in col_map]
    if not date_cols:
        return pd.DataFrame()

    wide = sub[date_cols].copy()
    wide.columns = [col_map[c].normalize() for c in date_cols]
    wide = wide.sort_index(axis=1)
    return wide


def _date_range_filter(columns: Sequence[pd.Timestamp], start: str | None, end: str | None) -> list[pd.Timestamp]:
    if not start and not end:
        return list(columns)
    s = pd.to_datetime(start).normalize() if start else None
    e = pd.to_datetime(end).normalize() if end else None
    out: list[pd.Timestamp] = []
    for c in columns:
        if s is not None and c < s:
            continue
        if e is not None and c > e:
            continue
        out.append(c)
    return out


def aggregate_one_file(path: Path, cfg: Config) -> pd.DataFrame:
    df = pq.read_table(str(path)).to_pandas()
    if "__sheet_name" not in df.columns or "__row_id" not in df.columns:
        raise ValueError(f"{path} missing __sheet_name/__row_id")

    price = _sheet_wide(df, cfg.price_sheet)
    premium = _sheet_wide(df, cfg.premium_sheet)
    status = _sheet_wide(df, cfg.status_sheet)
    if price.empty or premium.empty:
        return pd.DataFrame(columns=["date", "avg_conv_premium", "count"])

    ignore_status_from = pd.to_datetime(cfg.ignore_status_from).normalize() if cfg.ignore_status_from else None
    price_premium_dates = sorted(set(price.columns) & set(premium.columns))
    if ignore_status_from is None:
        if status.empty:
            return pd.DataFrame(columns=["date", "avg_conv_premium", "count"])
        common_dates = [d for d in price_premium_dates if d in status.columns]
    else:
        common_dates = [
            d
            for d in price_premium_dates
            if d >= ignore_status_from or (not status.empty and d in status.columns)
        ]
    common_dates = _date_range_filter(common_dates, cfg.start, cfg.end)
    if not common_dates:
        return pd.DataFrame(columns=["date", "avg_conv_premium", "count"])

    common_index = price.index.intersection(premium.index)
    price_num = _to_number(price.reindex(index=common_index, columns=common_dates))
    premium_num = _to_number(premium.reindex(index=common_index, columns=common_dates))

    trading_mask = pd.DataFrame(True, index=common_index, columns=common_dates)
    status_filter_dates = [
        d for d in common_dates if ignore_status_from is None or d < ignore_status_from
    ]
    if status_filter_dates:
        status_str = status.reindex(index=common_index, columns=status_filter_dates).astype("string")
        trading_mask.loc[:, status_filter_dates] = status_str.isin(list(cfg.trading_values))

    # ignore_status_from 之前按交易状态过滤；自该日期起忽略交易状态
    eligible = trading_mask & (price_num > cfg.price_threshold) & premium_num.notna()

    rows: list[dict] = []
    for d in common_dates:
        m = eligible[d]
        if bool(m.any()):
            rows.append(
                {
                    "date": d.date(),
                    "avg_conv_premium": float(premium_num.loc[m, d].mean()),
                    "count": int(m.sum()),
                }
            )
        else:
            rows.append({"date": d.date(), "avg_conv_premium": float("nan"), "count": 0})

    return pd.DataFrame(rows)


def compute_time_series(cfg: Config) -> pd.DataFrame:
    files = _iter_monthly_parquets(cfg.root, cfg.parquet_dir)
    if not files:
        raise RuntimeError("No monthly parquet files found.")

    parts: list[pd.DataFrame] = []
    for i, fp in enumerate(files, 1):
        part = aggregate_one_file(fp, cfg)
        if not part.empty:
            parts.append(part)
        if i % 24 == 0:
            print(f"processed {i}/{len(files)}: {fp.name}")

    if not parts:
        return pd.DataFrame(columns=["date", "avg_conv_premium", "count"])

    all_df = pd.concat(parts, ignore_index=True)
    date_counts = all_df.groupby("date", as_index=False).agg(count=("count", "sum"))

    valid = all_df[(all_df["count"] > 0) & all_df["avg_conv_premium"].notna()].copy()
    if valid.empty:
        out = date_counts.copy()
        out["avg_conv_premium"] = np.nan
        out["count"] = out["count"].astype("int64")
        return out.sort_values("date", kind="stable")[["date", "avg_conv_premium", "count"]].reset_index(drop=True)

    valid["w"] = valid["avg_conv_premium"] * valid["count"]
    wsum = valid.groupby("date", as_index=False).agg(w_sum=("w", "sum"), w_count=("count", "sum"))
    out = date_counts.merge(wsum, on="date", how="left")
    out["avg_conv_premium"] = out["w_sum"] / out["w_count"]
    out = out.drop(columns=["w_sum", "w_count"])
    out["count"] = out["count"].astype("int64")
    return out.sort_values("date", kind="stable")[["date", "avg_conv_premium", "count"]].reset_index(drop=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="以 Parquet 分片为数据源：计算 收盘价>阈值 的转债 转股溢价率(均值) 时间序列；并用 交易状态 剔除非交易日。"
    )
    parser.add_argument("--root", default=str(Path.cwd()), help="工作目录（默认当前目录）")
    parser.add_argument("--price-threshold", type=float, default=170.0, help="价格阈值（默认 170）")
    parser.add_argument("--trading-values", default="交易,正常交易", help="视为交易日的交易状态（逗号分隔）")
    parser.add_argument(
        "--ignore-status-from",
        default=None,
        help="自该日期起忽略交易状态过滤，格式 YYYY-MM-DD（可选）",
    )
    parser.add_argument("--start", default="2015-01-01", help="起始日期 YYYY-MM-DD（默认 2015-01-01）")
    parser.add_argument("--end", default=None, help="结束日期 YYYY-MM-DD（可选）")
    parser.add_argument(
        "--out",
        default="cb_price_gt170_avg_conv_premium_ts.xlsx",
        help="输出文件名（.xlsx 或 .csv；默认 xlsx）",
    )
    args = parser.parse_args()

    cfg = Config(
        root=Path(args.root).resolve(),
        price_threshold=float(args.price_threshold),
        trading_values=tuple(s.strip() for s in str(args.trading_values).split(",") if s.strip()),
        ignore_status_from=args.ignore_status_from,
        start=args.start,
        end=args.end,
        out=Path(args.out),
    )

    ts = compute_time_series(cfg)
    out_path = (cfg.root / cfg.out).resolve()
    if out_path.suffix.lower() == ".xlsx":
        with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
            ts.to_excel(writer, index=False, sheet_name="ts")
        print(f"[ok] wrote: {out_path} (rows={len(ts)})")
    elif out_path.suffix.lower() == ".csv":
        ts.to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"[ok] wrote: {out_path} (rows={len(ts)})")
    else:
        raise SystemExit("out 必须是 .xlsx 或 .csv")


if __name__ == "__main__":
    # 避免某些环境里 pandas 线程告警
    os.environ.setdefault("OMP_NUM_THREADS", "1")
    main()
