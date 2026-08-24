from __future__ import annotations

import argparse
import re
from pathlib import Path

import numpy as np
import pandas as pd


SHEETS = ("平价", "转股溢价率", "余额")
OUTPUT_COLUMN = "平价95-100转债_转股溢价率余额加权均值"


def monthly_parquet_files(parquet_root: Path) -> list[Path]:
    """只读取 yyyy/yyyymm.parquet 月度分片。"""
    files = [
        path
        for path in parquet_root.glob("*/*.parquet")
        if path.parent.name.isdigit() and re.fullmatch(r"\d{6}", path.stem)
    ]
    return sorted(files)


def date_columns(columns: pd.Index) -> dict[object, pd.Timestamp]:
    result: dict[object, pd.Timestamp] = {}
    for column in columns:
        if column in {"__sheet_name", "__row_id"}:
            continue
        date = pd.to_datetime(column, errors="coerce")
        if pd.notna(date):
            result[column] = pd.Timestamp(date).normalize()
    return result


def sheet_wide(frame: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    part = frame.loc[frame["__sheet_name"].eq(sheet_name)].copy()
    if part.empty:
        return pd.DataFrame()

    mapping = date_columns(part.columns)
    if not mapping:
        return pd.DataFrame()

    part = part[["__row_id", *mapping]].rename(columns=mapping)
    part["__row_id"] = part["__row_id"].astype("string")
    return part.set_index("__row_id")


def to_number(series: pd.Series) -> pd.Series:
    text = series.astype("string").str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(text, errors="coerce")


def aggregate_one_file(
    path: Path,
    lower: float = 95.0,
    upper: float = 100.0,
) -> pd.DataFrame:
    frame = pd.read_parquet(path)
    required = {"__sheet_name", "__row_id"}
    if not required.issubset(frame.columns):
        raise ValueError(f"Parquet 缺少必要列 {required}: {path}")

    wide = {name: sheet_wide(frame, name) for name in SHEETS}
    if any(item.empty for item in wide.values()):
        return pd.DataFrame()

    dates = sorted(set.intersection(*(set(item.columns) for item in wide.values())))
    rows: list[dict[str, object]] = []
    for date in dates:
        parity = to_number(wide["平价"][date])
        premium = to_number(wide["转股溢价率"][date])
        balance = to_number(wide["余额"][date])

        # 与日报平价分组的右端点口径一致：(95, 100]。
        valid = (
            parity.gt(lower)
            & parity.le(upper)
            & premium.notna()
            & balance.notna()
            & np.isfinite(parity)
            & np.isfinite(premium)
            & np.isfinite(balance)
            & balance.gt(0)
        )
        valid_balance = balance.loc[valid]
        balance_sum = float(valid_balance.sum())
        weighted_mean = (
            float((premium.loc[valid] * valid_balance).sum() / balance_sum)
            if balance_sum > 0
            else np.nan
        )
        rows.append(
            {
                "日期": date,
                "有效个券数": int(valid.sum()),
                "有效余额合计": balance_sum if balance_sum > 0 else np.nan,
                OUTPUT_COLUMN: weighted_mean,
                "来源分片": path.name,
            }
        )
    return pd.DataFrame(rows)


def build_series(
    parquet_root: Path,
    lower: float = 95.0,
    upper: float = 100.0,
    start: str | None = None,
    end: str | None = None,
) -> pd.DataFrame:
    files = monthly_parquet_files(parquet_root)
    if not files:
        raise FileNotFoundError(f"未找到月度 Parquet 分片: {parquet_root.resolve()}")

    start_date = pd.Timestamp(start).normalize() if start else None
    end_date = pd.Timestamp(end).normalize() if end else None
    parts: list[pd.DataFrame] = []
    for path in files:
        month = pd.Timestamp(path.stem + "01")
        if start_date is not None and month + pd.offsets.MonthEnd(0) < start_date:
            continue
        if end_date is not None and month > end_date:
            continue
        part = aggregate_one_file(path, lower=lower, upper=upper)
        if not part.empty:
            parts.append(part)

    if not parts:
        raise ValueError("指定日期范围内没有可计算数据")

    result = pd.concat(parts, ignore_index=True)
    if start_date is not None:
        result = result.loc[result["日期"].ge(start_date)]
    if end_date is not None:
        result = result.loc[result["日期"].le(end_date)]

    # 理论上月度分片无重叠；如补档造成日期重复，以较新的文件名为准。
    result = result.sort_values(["日期", "来源分片"])
    result = result.drop_duplicates("日期", keep="last").sort_values("日期")
    return result.reset_index(drop=True)


def write_outputs(
    result: pd.DataFrame,
    output: Path,
    parquet_root: Path,
    lower: float,
    upper: float,
) -> tuple[Path, Path]:
    output.parent.mkdir(parents=True, exist_ok=True)
    csv_path = output.with_suffix(".csv")
    info = pd.DataFrame(
        [
            {"项目": "数据源", "说明": str(parquet_root.resolve())},
            {"项目": "筛选范围", "说明": f"{lower} < 平价 <= {upper}"},
            {"项目": "计算公式", "说明": "sum(转股溢价率 × 余额) / sum(余额)"},
            {"项目": "有效样本", "说明": "三字段均有效且余额>0"},
            {
                "项目": "日期范围",
                "说明": f"{result['日期'].min():%Y-%m-%d} ~ {result['日期'].max():%Y-%m-%d}",
            },
        ]
    )
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        info.to_excel(writer, sheet_name="说明", index=False)
        result.to_excel(writer, sheet_name="时间序列", index=False)
    result.to_csv(csv_path, index=False, encoding="utf-8-sig")
    return output.resolve(), csv_path.resolve()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="以月度 Parquet 为数据源，计算平价区间内转债的转股溢价率余额加权均值。"
    )
    parser.add_argument("--parquet-root", default="data/转债个券历史序列")
    parser.add_argument("--lower", type=float, default=95.0, help="平价下界（不含）")
    parser.add_argument("--upper", type=float, default=100.0, help="平价上界（含）")
    parser.add_argument("--start", help="起始日期 YYYY-MM-DD")
    parser.add_argument("--end", help="结束日期 YYYY-MM-DD")
    parser.add_argument(
        "--output",
        default="outputs/平价95-100转股溢价率余额加权均值.xlsx",
    )
    args = parser.parse_args()
    if args.lower >= args.upper:
        parser.error("lower 必须小于 upper")

    root = Path(args.parquet_root)
    result = build_series(
        root,
        lower=args.lower,
        upper=args.upper,
        start=args.start,
        end=args.end,
    )
    xlsx, csv = write_outputs(result, Path(args.output), root, args.lower, args.upper)
    latest = result.iloc[-1]
    print(f"[ok] Excel: {xlsx}")
    print(f"[ok] CSV:   {csv}")
    print(
        f"[latest] {latest['日期']:%Y-%m-%d} | 个券={latest['有效个券数']} | "
        f"余额={latest['有效余额合计']:.4f} | 加权均值={latest[OUTPUT_COLUMN]:.6f}"
    )


if __name__ == "__main__":
    main()
