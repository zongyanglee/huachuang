from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple

import numpy as np
import pandas as pd
from scipy.optimize import curve_fit

sys.path.append(str(Path(__file__).resolve().parents[1]))

from 底稿更新 import load_original_data


@dataclass(frozen=True)
class RunConfig:
    parquet_root: str = "data/转债个券历史序列"
    output_xlsx: str = "剩余期限周分档_delta_parquet.xlsx"

    max_weeks: int = 311
    target_x: float = 100.0
    min_fit_points: int = 8

    plain_low: float = 70.0
    plain_high: float = 130.0
    winsor_low_q: float = 0.03
    winsor_high_q: float = 0.97


def inverse_cubic(x: np.ndarray | float, a: float, b: float, c: float, d: float) -> np.ndarray | float:
    return a / np.power(x, 3) + b / np.power(x, 2) + c / x + d


def delta_cal(x: np.ndarray | float, a: float, b: float, c: float, d: float) -> np.ndarray | float:
    return (-2 * a / np.power(x, 3) - b / np.power(x, 2) + d + 100) / 100


def _date_cols(df: pd.DataFrame) -> List[pd.Timestamp]:
    return sorted([c for c in df.columns if isinstance(c, pd.Timestamp)])


def _winsorize_by_quantile(df: pd.DataFrame, col: str, low_q: float, high_q: float) -> pd.DataFrame:
    if df.empty or col not in df.columns:
        return df
    try:
        low = np.nanquantile(df[col].values, low_q)
        high = np.nanquantile(df[col].values, high_q)
        return df[(df[col] > low) & (df[col] < high)].dropna(axis=0)
    except Exception:
        return df


def _fit_detail(sub: pd.DataFrame, target_x: float, min_points: int) -> dict:
    work = sub.copy()
    work = work.replace("", np.nan).replace(0, np.nan).dropna(subset=["平价", "转股溢价率"])
    if len(work) < min_points:
        return {"a": np.nan, "b": np.nan, "c": np.nan, "d": np.nan, "delta": np.nan, "n": int(len(work))}
    x = work["平价"].astype(float).values
    y = work["转股溢价率"].astype(float).values
    try:
        popt, _ = curve_fit(inverse_cubic, x, y, maxfev=20000)
        a, b, c, d = (float(v) for v in popt)
        return {
            "a": a,
            "b": b,
            "c": c,
            "d": d,
            "delta": float(delta_cal(float(target_x), a, b, c, d)),
            "n": int(len(work)),
        }
    except Exception:
        return {"a": np.nan, "b": np.nan, "c": np.nan, "d": np.nan, "delta": np.nan, "n": int(len(work))}


def main(config: RunConfig) -> Path:
    data = load_original_data(source_type="parquet", parquet_root=config.parquet_root, force_refresh=False)
    required = ["平价", "转股溢价率", "剩余期限"]
    missing = [s for s in required if s not in data]
    if missing:
        raise KeyError(f"parquet 数据缺少 sheet: {missing}")

    plain_df = data["平价"]
    premium_df = data["转股溢价率"]
    remain_year_df = data["剩余期限"]

    date_sets = [set(_date_cols(d)) for d in (plain_df, premium_df, remain_year_df)]
    date_range = sorted(set.intersection(*date_sets))
    if not date_range:
        raise ValueError("未找到可用日期列交集。")

    remain_week_df = remain_year_df.replace("", 0).apply(pd.to_numeric, errors="coerce") * 52

    df_data = pd.DataFrame(
        {
            "平价": plain_df[date_range].values.flatten(),
            "转股溢价率": premium_df[date_range].values.flatten(),
            "剩余期限(周)": remain_week_df[date_range].values.flatten(),
        }
    )
    df_data = df_data.replace("", np.nan).dropna(axis=0, how="any")
    for col in df_data.columns:
        df_data[col] = pd.to_numeric(df_data[col], errors="coerce")
    df_data = df_data.dropna(axis=0, how="any")

    # 过滤条件（与周分档截面拟合保持一致：仅平价区间）
    df_data = df_data[(df_data["平价"] > config.plain_low) & (df_data["平价"] < config.plain_high)]

    rows: List[dict] = []
    for i in range(0, config.max_weeks):
        lo, hi = i, i + 1
        sub = df_data[(df_data["剩余期限(周)"] >= lo) & (df_data["剩余期限(周)"] < hi)].copy()
        sub = _winsorize_by_quantile(sub, col="转股溢价率", low_q=config.winsor_low_q, high_q=config.winsor_high_q)
        detail = _fit_detail(sub.rename(columns={"剩余期限(周)": "剩余期限(周)"}), target_x=config.target_x, min_points=config.min_fit_points)
        rows.append(
            {
                "剩余期限(周)": f"{lo}-{hi}",
                "N": detail["n"],
                "a": detail["a"],
                "b": detail["b"],
                "c": detail["c"],
                "d": detail["d"],
                "delta(平价=100)": detail["delta"],
            }
        )

    out_df = pd.DataFrame(rows, columns=["剩余期限(周)", "N", "a", "b", "c", "d", "delta(平价=100)"])

    out_path = Path(config.output_xlsx).resolve()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        info = pd.DataFrame(
            [
                {"key": "数据源", "value": f"parquet: {Path(config.parquet_root).resolve()}"},
                {"key": "口径", "value": "参照分组拟合溢价率：剩余期限(年)*52 -> 周分档[lo,hi)做全样本截面拟合"},
                {"key": "过滤", "value": f"平价({config.plain_low},{config.plain_high})"},
                {"key": "winsorize", "value": f"转股溢价率分位({config.winsor_low_q},{config.winsor_high_q})"},
                {"key": "max_weeks", "value": config.max_weeks},
                {"key": "delta", "value": f"在平价={config.target_x}处的导数 delta_cal"},
                {"key": "日期范围(用于拼截面)", "value": f"{date_range[0].date()} ~ {date_range[-1].date()} (n={len(date_range)})"},
            ]
        )
        info.to_excel(writer, sheet_name="说明", index=False)
        out_df.to_excel(writer, sheet_name="delta_周分档", index=False)

    return out_path


if __name__ == "__main__":
    out = main(RunConfig())
    print(f"[ok] wrote: {out}")

