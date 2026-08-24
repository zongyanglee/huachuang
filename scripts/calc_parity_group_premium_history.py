from __future__ import annotations

import json
import re
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
PARQUET_ROOT = ROOT / "data/转债个券历史序列"
OUTPUT_DIR = ROOT / "outputs" / "parity_group_premium_history"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SHEETS = ("平价", "转股溢价率", "余额")
GROUPS = [
    ("70-90", 70.0, 90.0),
    ("90-110", 90.0, 110.0),
    ("110-130", 110.0, 130.0),
    ("130-150", 130.0, 150.0),
]


def monthly_parquet_files() -> list[Path]:
    files = [
        path
        for path in PARQUET_ROOT.glob("*/*.parquet")
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


def weighted_mean(values: pd.Series, weights: pd.Series) -> float:
    valid = values.notna() & weights.notna() & weights.gt(0)
    if not valid.any():
        return np.nan
    return float(np.average(values.loc[valid], weights=weights.loc[valid]))


def aggregate_one_file(path: Path) -> list[dict[str, object]]:
    frame = pd.read_parquet(path)
    required = {"__sheet_name", "__row_id"}
    if not required.issubset(frame.columns):
        raise ValueError(f"Parquet缺少必要列 {required}: {path}")

    wide = {name: sheet_wide(frame, name) for name in SHEETS}
    if any(item.empty for item in wide.values()):
        return []

    dates = sorted(set.intersection(*(set(item.columns) for item in wide.values())))
    rows: list[dict[str, object]] = []
    for date in dates:
        parity = to_number(wide["平价"][date])
        premium = to_number(wide["转股溢价率"][date])
        balance = to_number(wide["余额"][date])

        for label, lower, upper in GROUPS:
            valid = (
                parity.gt(lower)
                & parity.le(upper)
                & premium.notna()
                & np.isfinite(parity)
                & np.isfinite(premium)
            )
            premium_sub = premium.loc[valid]
            balance_sub = balance.loc[valid]
            rows.append(
                {
                    "日期": date.strftime("%Y-%m-%d"),
                    "平价区间": label,
                    "样本数": int(premium_sub.shape[0]),
                    "余额合计(亿元)": float(balance_sub.dropna().sum()) if not balance_sub.empty else np.nan,
                    "算术平均转股溢价率(%)": float(premium_sub.mean()) if not premium_sub.empty else np.nan,
                    "余额加权平均转股溢价率(%)": weighted_mean(premium_sub, balance_sub),
                    "中位数转股溢价率(%)": float(premium_sub.median()) if not premium_sub.empty else np.nan,
                    "来源分片": path.name,
                }
            )
    return rows


def pivot_metric(history: pd.DataFrame, value_col: str) -> pd.DataFrame:
    out = history.pivot(index="日期", columns="平价区间", values=value_col)
    return out.reindex(columns=[label for label, _, _ in GROUPS]).reset_index()


def main() -> None:
    files = monthly_parquet_files()
    if not files:
        raise FileNotFoundError(f"未找到月度 parquet 分片: {PARQUET_ROOT}")

    rows: list[dict[str, object]] = []
    for path in files:
        rows.extend(aggregate_one_file(path))

    history = pd.DataFrame(rows)
    if history.empty:
        raise ValueError("未计算出有效历史序列")

    history = history.sort_values(["日期", "平价区间", "来源分片"]).drop_duplicates(
        ["日期", "平价区间"], keep="last"
    )
    history = history.reset_index(drop=True)

    payload = {
        "source_root": str(PARQUET_ROOT.relative_to(ROOT)),
        "source_files": [str(path.relative_to(ROOT)) for path in files],
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        "date_start": str(history["日期"].min()),
        "date_end": str(history["日期"].max()),
        "trading_days": int(history["日期"].nunique()),
        "group_rule": "按平价分组，区间为左开右闭：(70,90]、(90,110]、(110,130]、(130,150]。",
        "metric_rule": "算术平均和中位数使用转股溢价率有效样本；余额加权平均使用余额>0且转股溢价率有效样本，公式为 sum(转股溢价率*余额)/sum(余额)。转股溢价率单位为百分点。",
        "groups": [label for label, _, _ in GROUPS],
        "latest_rows": history[history["日期"].eq(history["日期"].max())]
        .replace({np.nan: None})
        .to_dict(orient="records"),
    }

    json_path = OUTPUT_DIR / "parity_group_premium_history_meta.json"
    csv_path = OUTPUT_DIR / "平价分组转股溢价率历史序列_长表.csv"
    arithmetic_path = OUTPUT_DIR / "算术平均转股溢价率_宽表.csv"
    weighted_path = OUTPUT_DIR / "余额加权平均转股溢价率_宽表.csv"
    median_path = OUTPUT_DIR / "中位数转股溢价率_宽表.csv"
    count_path = OUTPUT_DIR / "样本数_宽表.csv"

    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    history.to_csv(csv_path, index=False, encoding="utf-8-sig")
    pivot_metric(history, "算术平均转股溢价率(%)").to_csv(arithmetic_path, index=False, encoding="utf-8-sig")
    pivot_metric(history, "余额加权平均转股溢价率(%)").to_csv(weighted_path, index=False, encoding="utf-8-sig")
    pivot_metric(history, "中位数转股溢价率(%)").to_csv(median_path, index=False, encoding="utf-8-sig")
    pivot_metric(history, "样本数").to_csv(count_path, index=False, encoding="utf-8-sig")

    print(json_path)
    print(csv_path)
    print("files", len(files), "days", payload["trading_days"], payload["date_start"], payload["date_end"])
    print(history[history["日期"].eq(history["日期"].max())].to_string(index=False))


if __name__ == "__main__":
    main()
