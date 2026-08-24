from __future__ import annotations

import argparse
import importlib.util
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Sequence, Tuple

import numpy as np
import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from 底稿更新 import load_original_data


@dataclass(frozen=True)
class RunConfig:
    parquet_root: str = "data/转债个券历史序列"
    output_xlsx: str = "剩余期限0-2与2-6百元溢价率_parquet.xlsx"
    start_date: str = "2018-01-01"
    window_size: int = 5
    plain_low: float = 70.0
    plain_high: float = 130.0
    turnover_max: float = 50.0
    target_x: float = 100.0


def _load_group_fit_module():
    module_path = Path(__file__).resolve().parents[1] / "【计算】分组拟合溢价率.py"
    spec = importlib.util.spec_from_file_location("group_fit_premium", module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载分组拟合脚本: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _date_cols(df: pd.DataFrame) -> List[pd.Timestamp]:
    return sorted([c for c in df.columns if isinstance(c, pd.Timestamp)])


def _build_window_df(
    plain_df: pd.DataFrame,
    premium_df: pd.DataFrame,
    turnover_df: pd.DataFrame,
    remain_year_df: pd.DataFrame,
    cols: Sequence[pd.Timestamp],
    pure_bond_premium_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    data = {
        "平价": plain_df[list(cols)].values.flatten(),
        "转股溢价率": premium_df[list(cols)].values.flatten(),
        "换手率": turnover_df[list(cols)].values.flatten(),
        "剩余期限": remain_year_df[list(cols)].values.flatten(),
    }
    if pure_bond_premium_df is not None:
        data["纯债溢价率"] = pure_bond_premium_df[list(cols)].values.flatten()
    return pd.DataFrame(data)


def _prepare_date_range(dataframes: Sequence[pd.DataFrame], start_date: str, window_size: int) -> List[pd.Timestamp]:
    date_sets = [set(_date_cols(df)) for df in dataframes]
    dates = sorted(set.intersection(*date_sets))
    start_ts = pd.Timestamp(start_date)
    dates = [d for d in dates if d >= start_ts]
    if len(dates) < window_size:
        raise ValueError(f"可用日期列不足：n_dates={len(dates)}, window_size={window_size}, start_date={start_date}")
    return dates


def build_result(config: RunConfig) -> pd.DataFrame:
    group_fit = _load_group_fit_module()
    fit_premium_at_x = group_fit._fit_premium_at_x
    winsorize_by_quantile = group_fit._winsorize_by_quantile

    data = load_original_data(source_type="parquet", parquet_root=config.parquet_root, force_refresh=False)
    required = ["平价", "转股溢价率", "换手率", "剩余期限", "纯债溢价率"]
    missing = [sheet for sheet in required if sheet not in data]
    if missing:
        raise KeyError(f"parquet 数据缺少 sheet: {missing}")

    plain_df = data["平价"]
    premium_df = data["转股溢价率"]
    turnover_df = data["换手率"]
    remain_year_df = data["剩余期限"]
    pure_bond_premium_df = data["纯债溢价率"]

    dates = _prepare_date_range(
        [plain_df, premium_df, turnover_df, remain_year_df, pure_bond_premium_df],
        start_date=config.start_date,
        window_size=config.window_size,
    )

    fit_labels = ["0-2年", "2-6年"]
    labels = [
        *fit_labels,
        "0-2年纯债溢价率均值",
        "0-2年纯债溢价率中位数",
        "2-6年纯债溢价率均值",
        "2-6年纯债溢价率中位数",
    ]
    bins: List[Tuple[float, float]] = [(0.0, 2.0), (2.0, 6.0)]

    rows: List[dict] = []
    for i in range(config.window_size - 1, len(dates)):
        cols = dates[i - config.window_size + 1 : i + 1]
        base = _build_window_df(
            plain_df,
            premium_df,
            turnover_df,
            remain_year_df,
            cols,
            pure_bond_premium_df=pure_bond_premium_df,
        )
        base = base.replace("", np.nan)
        for col in ("平价", "转股溢价率", "换手率", "剩余期限", "纯债溢价率"):
            base[col] = pd.to_numeric(base[col], errors="coerce")
        base = base.dropna(subset=["平价", "转股溢价率", "换手率", "剩余期限"])
        base = base[
            (base["平价"] > config.plain_low)
            & (base["平价"] < config.plain_high)
            & (base["换手率"] < config.turnover_max)
        ]

        row = {"日期": dates[i]}
        for label, (lo, hi) in zip(fit_labels, bins):
            sub = base[(base["剩余期限"] > lo) & (base["剩余期限"] < hi)].copy()
            sub = winsorize_by_quantile(sub)
            row[label] = fit_premium_at_x(sub, config.target_x)

        daily_pure = pd.DataFrame(
            {
                "剩余期限": remain_year_df[dates[i]].values,
                "纯债溢价率": pure_bond_premium_df[dates[i]].values,
            }
        ).replace("", np.nan)
        for col in ("剩余期限", "纯债溢价率"):
            daily_pure[col] = pd.to_numeric(daily_pure[col], errors="coerce")
        daily_pure = daily_pure.dropna(subset=["剩余期限", "纯债溢价率"])
        for label, (lo, hi) in [
            ("0-2年纯债溢价率均值", (0.0, 2.0)),
            ("2-6年纯债溢价率均值", (2.0, 6.0)),
        ]:
            pure_sub = daily_pure[(daily_pure["剩余期限"] > lo) & (daily_pure["剩余期限"] < hi)]
            row[label] = float(pure_sub["纯债溢价率"].mean())
            row[label.replace("均值", "中位数")] = float(pure_sub["纯债溢价率"].median())
        rows.append(row)

    return pd.DataFrame(rows).set_index("日期").sort_index().reindex(columns=labels)


def write_excel(result: pd.DataFrame, config: RunConfig) -> Path:
    out_path = Path(config.output_xlsx).resolve()
    info = pd.DataFrame(
        [
            {"key": "数据源", "value": f"parquet: {Path(config.parquet_root).resolve()}"},
            {"key": "算法口径", "value": "复用【计算】分组拟合溢价率.py 的 _winsorize_by_quantile 与 _fit_premium_at_x"},
            {"key": "分组", "value": "剩余期限 0-2年、2-6年；边界口径为 >lo 且 <hi"},
            {"key": "窗口", "value": f"{config.window_size}日滚动窗口"},
            {"key": "过滤", "value": f"平价({config.plain_low},{config.plain_high}), 换手率<{config.turnover_max}"},
            {"key": "拟合目标", "value": f"平价={config.target_x} 的拟合转股溢价率"},
            {"key": "新增列", "value": "0-2年、2-6年纯债溢价率均值/中位数：当日截面中对应剩余期限分组的纯债溢价率统计值，不使用滚动窗口"},
            {"key": "起始日期", "value": config.start_date},
            {"key": "输出日期范围", "value": f"{result.index.min().date()} ~ {result.index.max().date()} (n={len(result)})"},
        ]
    )
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        info.to_excel(writer, sheet_name="说明", index=False)
        result.to_excel(writer, sheet_name="剩余期限百元溢价率")
    return out_path


def main(config: RunConfig) -> Path:
    result = build_result(config)
    return write_excel(result, config)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="使用 parquet 历史数据计算剩余期限 0-2年、2-6年的百元拟合溢价率")
    parser.add_argument("--parquet-root", default=RunConfig.parquet_root)
    parser.add_argument("--output-xlsx", default=RunConfig.output_xlsx)
    parser.add_argument("--start-date", default=RunConfig.start_date)
    args = parser.parse_args()

    out = main(
        RunConfig(
            parquet_root=args.parquet_root,
            output_xlsx=args.output_xlsx,
            start_date=args.start_date,
        )
    )
    print(f"[ok] wrote: {out}")
