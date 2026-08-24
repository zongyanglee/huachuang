from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

import numpy as np
import pandas as pd
from scipy.optimize import curve_fit

sys.path.append(str(Path(__file__).resolve().parents[1]))

from 底稿更新 import load_original_data


@dataclass(frozen=True)
class RunConfig:
    parquet_root: str = "data/转债个券历史序列"
    output_xlsx: str = "剩余期限分组_delta时间序列_parquet.xlsx"

    # 复用“分组拟合溢价率.py”的口径
    window_size: int = 5
    min_fit_points: int = 8
    winsor_low_q: float = 0.03
    winsor_high_q: float = 0.97

    # 过滤条件（与拟合脚本保持一致）
    plain_low: float = 70.0
    plain_high: float = 130.0
    turnover_max: float = 50.0

    target_x: float = 100.0


def inverse_cubic(x: np.ndarray | float, a: float, b: float, c: float, d: float) -> np.ndarray | float:
    return a / np.power(x, 3) + b / np.power(x, 2) + c / x + d


def delta_cal(x: np.ndarray | float, a: float, b: float, c: float, d: float) -> np.ndarray | float:
    # price = (premium + 100) / 100 * plain，对plain求导
    return (-2 * a / np.power(x, 3) - b / np.power(x, 2) + d + 100) / 100


def _date_cols(df: pd.DataFrame) -> List[pd.Timestamp]:
    return sorted([c for c in df.columns if isinstance(c, pd.Timestamp)])


def _as_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def _winsorize_by_quantile(df: pd.DataFrame, col: str, low_q: float, high_q: float) -> pd.DataFrame:
    if df.empty or col not in df.columns:
        return df
    try:
        low = np.nanquantile(df[col].values, low_q)
        high = np.nanquantile(df[col].values, high_q)
        return df[(df[col] > low) & (df[col] < high)].dropna(axis=0)
    except Exception:
        return df


def _build_window_df(
    plain_df: pd.DataFrame,
    premium_df: pd.DataFrame,
    turnover_df: pd.DataFrame,
    remain_year_df: pd.DataFrame,
    cols: Sequence[pd.Timestamp],
) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "平价": plain_df[list(cols)].values.flatten(),
            "转股溢价率": premium_df[list(cols)].values.flatten(),
            "换手率": turnover_df[list(cols)].values.flatten(),
            "剩余期限": remain_year_df[list(cols)].values.flatten(),
        }
    )


def _fit_delta_at_x(df_subset: pd.DataFrame, target_x: float, min_points: int) -> float:
    if df_subset is None or df_subset.empty:
        return float("nan")
    work = df_subset.copy()
    work = work.replace("", np.nan).replace(0, np.nan).dropna(subset=["平价", "转股溢价率"])
    if len(work) < min_points:
        return float("nan")
    x = work["平价"].astype(float).values
    y = work["转股溢价率"].astype(float).values
    try:
        popt, _ = curve_fit(inverse_cubic, x, y, maxfev=20000)
        a, b, c, d = (float(v) for v in popt)
        return float(delta_cal(float(target_x), a, b, c, d))
    except Exception:
        return float("nan")


def main(config: RunConfig) -> Path:
    data = load_original_data(source_type="parquet", parquet_root=config.parquet_root, force_refresh=False)

    required = ["平价", "转股溢价率", "换手率", "剩余期限"]
    missing = [s for s in required if s not in data]
    if missing:
        raise KeyError(f"parquet 数据缺少 sheet: {missing}")

    plain_df = data["平价"]
    premium_df = data["转股溢价率"]
    turnover_df = data["换手率"]
    remain_year_df = data["剩余期限"]

    # 日期交集，避免列缺失
    date_sets = [set(_date_cols(d)) for d in (plain_df, premium_df, turnover_df, remain_year_df)]
    dates = sorted(set.intersection(*date_sets))
    if len(dates) < config.window_size:
        raise ValueError(f"可用日期列不足以滚动窗口计算：n_dates={len(dates)}, window={config.window_size}")

    labels = ["0-1", "1-2", "2-3", "3-4", "4-5", "5-6"]
    bins: List[Tuple[float, float]] = [(0, 1), (1, 2), (2, 3), (3, 4), (4, 5), (5, 6)]

    rows: List[dict] = []
    for i in range(config.window_size - 1, len(dates)):
        cols = dates[i - config.window_size + 1 : i + 1]
        base = _build_window_df(plain_df, premium_df, turnover_df, remain_year_df, cols)

        for c in ("平价", "转股溢价率", "换手率", "剩余期限"):
            base[c] = _as_num(base[c])
        base = base.dropna(subset=["平价", "转股溢价率", "换手率", "剩余期限"])

        # 过滤条件（同拟合脚本）
        base = base[
            (base["平价"] > config.plain_low)
            & (base["平价"] < config.plain_high)
            & (base["换手率"] < config.turnover_max)
        ]

        row = {"日期": dates[i]}
        for lb, (lo, hi) in zip(labels, bins):
            sub = base[(base["剩余期限"] > lo) & (base["剩余期限"] < hi)].copy()
            sub = _winsorize_by_quantile(sub, col="转股溢价率", low_q=config.winsor_low_q, high_q=config.winsor_high_q)
            row[lb] = _fit_delta_at_x(sub, target_x=config.target_x, min_points=config.min_fit_points)
        rows.append(row)

    out_df = pd.DataFrame(rows).set_index("日期").sort_index()
    out_df = out_df.reindex(columns=labels)

    out_path = Path(config.output_xlsx).resolve()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        info = pd.DataFrame(
            [
                {"key": "数据源", "value": f"parquet: {Path(config.parquet_root).resolve()}"},
                {"key": "分组", "value": "剩余期限(0-1,1-2,2-3,3-4,4-5,5-6)"},
                {"key": "窗口", "value": f"{config.window_size}日滚动窗口（逐日拟合）"},
                {"key": "过滤", "value": f"平价({config.plain_low},{config.plain_high}), 换手率<{config.turnover_max}"},
                {"key": "winsorize", "value": f"转股溢价率分位({config.winsor_low_q},{config.winsor_high_q})"},
                {"key": "拟合函数", "value": "inverse_cubic(premium ~ plain)"},
                {"key": "delta", "value": f"在平价={config.target_x}处的导数 delta_cal"},
                {"key": "日期范围", "value": f"{out_df.index.min().date()} ~ {out_df.index.max().date()} (n={len(out_df)})"},
            ]
        )
        info.to_excel(writer, sheet_name="说明", index=False)
        out_df.to_excel(writer, sheet_name="delta_剩余期限", index=True)

    return out_path


if __name__ == "__main__":
    out = main(RunConfig())
    print(f"[ok] wrote: {out}")

