from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from 底稿更新 import load_original_data


@dataclass(frozen=True)
class RunConfig:
    parquet_root: str = "data/转债个券历史序列"
    output_xlsx: str = "分组波动率中位数序列_parquet.xlsx"

    # 正股波动率口径：默认 20 日
    stock_vol_sheet: str = "正股20日波动率"
    # 转债隐含波动率
    iv_sheet: str = "隐含波动率"


def _date_cols(df: pd.DataFrame) -> List[pd.Timestamp]:
    return sorted([c for c in df.columns if isinstance(c, pd.Timestamp)])


def _as_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def _bucket_left_closed_right_open(x: pd.Series, bins: Sequence[Tuple[float, float]], labels: Sequence[str]) -> pd.Series:
    xv = _as_num(x)
    out = pd.Series(np.nan, index=x.index, dtype="object")
    for (lo, hi), lb in zip(bins, labels):
        out[(xv >= lo) & (xv < hi)] = lb
    return out


def _median_by_group(metric: pd.Series, group_label: pd.Series, label_order: Sequence[str]) -> Dict[str, float]:
    metric_num = _as_num(metric)
    out: Dict[str, float] = {}
    for lb in label_order:
        mask = (group_label == lb) & metric_num.notna()
        if not mask.any():
            out[lb] = float("nan")
        else:
            out[lb] = float(metric_num[mask].median())
    return out


def _sector_label_from_industry(industry_s: pd.Series) -> pd.Series:
    sector_map: Dict[str, List[str]] = {
        "科技": ["传媒", "电子", "国防军工", "计算机", "通信"],
        "金融": ["非银金融", "银行"],
        "制造": ["电力设备", "机械设备", "汽车", "轻工制造"],
        "消费": ["农林牧渔", "纺织服饰", "家用电器", "商贸零售", "社会服务", "食品饮料", "医药生物", "美容护理"],
        "周期": ["基础化工", "钢铁", "公用事业", "环保", "建筑材料", "建筑装饰", "交通运输", "煤炭", "石油石化", "有色金属"],
    }
    out = pd.Series(np.nan, index=industry_s.index, dtype="object")
    for sec, inds in sector_map.items():
        out[industry_s.isin(inds)] = sec
    return out


def _rating_group_label(rating_raw: pd.Series) -> pd.Series:
    r = rating_raw.astype("object")
    invalid = r.isna() | r.astype(str).str.strip().isin({"0", "0.0", ""})
    r = r.mask(invalid, np.nan)
    out = pd.Series(np.nan, index=r.index, dtype="object")
    out[r.isin(["AAA", "AA+"])] = "AAA/AA+"
    out[r.isin(["AA", "AA-"])] = "AA/AA-"
    out[r.isin(["A+", "A"])] = "A/A-"
    return out


def _stock_bond_label(stock_bond_raw: pd.Series) -> pd.Series:
    sb = _as_num(stock_bond_raw)
    out = pd.Series(np.nan, index=sb.index, dtype="object")
    out[sb < -20] = "偏债型"
    out[(sb >= -20) & (sb < 20)] = "平衡型"
    out[sb >= 20] = "偏股型"
    return out


def _calc_one_metric_all_groups(
    metric_df: pd.DataFrame,
    metric_name: str,
    parity_df: pd.DataFrame,
    stock_bond_df: pd.DataFrame,
    rating_df: pd.DataFrame,
    remain_year_df: pd.DataFrame,
    balance_df: pd.DataFrame,
    stock_value_df: pd.DataFrame,
    sector_label: pd.Series,
    dates: List[pd.Timestamp],
) -> Dict[str, pd.DataFrame]:
    out_tables: Dict[str, pd.DataFrame] = {}

    group_specs = [
        ("板块", ["科技", "金融", "制造", "消费", "周期"]),
        ("股债型", ["偏债型", "平衡型", "偏股型"]),
        ("平价", ["70-90", "90-110", "110-130", "130-150"]),
        ("债项评级", ["AAA/AA+", "AA/AA-", "A/A-"]),
        ("剩余期限", ["0-1", "1-2", "2-3", "3-4", "4-5", "5-6"]),
        ("余额", ["0-3", "3-10", "10-20", "20-50", "50+"]),
        ("正股市值", ["0-50", "50-300", "300+"]),
    ]

    frames: Dict[str, List[dict]] = {name: [] for name, _ in group_specs}
    idx = metric_df.index

    for d in dates:
        metric_s = metric_df[d]

        # 分组均以当日口径动态生成（同前规则；左闭右开；不加过滤）
        sb_label = _stock_bond_label(stock_bond_df[d])
        plain_label = _bucket_left_closed_right_open(_as_num(parity_df[d]), [(70, 90), (90, 110), (110, 130), (130, 150)], ["70-90", "90-110", "110-130", "130-150"])
        rating_label = _rating_group_label(rating_df[d])
        ry_label = _bucket_left_closed_right_open(_as_num(remain_year_df[d]), [(0, 1), (1, 2), (2, 3), (3, 4), (4, 5), (5, 6)], ["0-1", "1-2", "2-3", "3-4", "4-5", "5-6"])

        bal0 = _as_num(balance_df[d])
        bal_label = _bucket_left_closed_right_open(bal0, [(0, 3), (3, 10), (10, 20), (20, 50)], ["0-3", "3-10", "10-20", "20-50"])
        bal_label[bal0 >= 50] = "50+"

        sv0 = _as_num(stock_value_df[d])
        sv_label = _bucket_left_closed_right_open(sv0, [(0, 50), (50, 300)], ["0-50", "50-300"])
        sv_label[sv0 >= 300] = "300+"

        labels_by_group = {
            "板块": sector_label,
            "股债型": sb_label,
            "平价": plain_label,
            "债项评级": rating_label,
            "剩余期限": ry_label,
            "余额": bal_label,
            "正股市值": sv_label,
        }

        for group_name, label_order in group_specs:
            g_label = labels_by_group[group_name]
            row = {"日期": d}
            row.update(_median_by_group(metric_s, g_label, label_order))
            frames[group_name].append(row)

    for group_name, label_order in group_specs:
        df = pd.DataFrame(frames[group_name]).set_index("日期").sort_index()
        df = df.reindex(columns=list(label_order))
        out_tables[f"{metric_name}_{group_name}"] = df

    return out_tables


def main(config: RunConfig) -> Path:
    data = load_original_data(source_type="parquet", parquet_root=config.parquet_root, force_refresh=False)

    required = [
        config.iv_sheet,
        config.stock_vol_sheet,
        "平价",
        "平价底价溢价率",
        "债项评级",
        "剩余期限",
        "余额",
        "正股市值",
        "总表",
    ]
    missing = [s for s in required if s not in data]
    if missing:
        raise KeyError(f"parquet 数据缺少 sheet: {missing}")

    iv_df = data[config.iv_sheet]
    stock_vol_df = data[config.stock_vol_sheet]
    parity_df = data["平价"]
    stock_bond_df = data["平价底价溢价率"]
    rating_df = data["债项评级"]
    remain_year_df = data["剩余期限"]
    balance_df = data["余额"]
    stock_value_df = data["正股市值"]

    # 申万行业来自总表（静态字段）
    total_df = data["总表"].copy()
    if "__row_id" in total_df.columns:
        total_df = total_df.set_index("__row_id")
    industry_s = total_df.get("申万行业")
    if industry_s is None:
        raise KeyError("`总表` 中未找到 `申万行业` 字段。")
    industry_s = industry_s.reindex(iv_df.index)
    sector_label = _sector_label_from_industry(industry_s)

    # 日期范围：取所有参与计算的 sheet 的日期交集
    date_sets = [
        set(_date_cols(iv_df)),
        set(_date_cols(stock_vol_df)),
        set(_date_cols(parity_df)),
        set(_date_cols(stock_bond_df)),
        set(_date_cols(rating_df)),
        set(_date_cols(remain_year_df)),
        set(_date_cols(balance_df)),
        set(_date_cols(stock_value_df)),
    ]
    dates = sorted(set.intersection(*date_sets))
    if not dates:
        raise ValueError("无法获得日期列交集，无法计算。")

    tables: Dict[str, pd.DataFrame] = {}
    tables.update(
        _calc_one_metric_all_groups(
            metric_df=stock_vol_df,
            metric_name=config.stock_vol_sheet,
            parity_df=parity_df,
            stock_bond_df=stock_bond_df,
            rating_df=rating_df,
            remain_year_df=remain_year_df,
            balance_df=balance_df,
            stock_value_df=stock_value_df,
            sector_label=sector_label,
            dates=dates,
        )
    )
    tables.update(
        _calc_one_metric_all_groups(
            metric_df=iv_df,
            metric_name=config.iv_sheet,
            parity_df=parity_df,
            stock_bond_df=stock_bond_df,
            rating_df=rating_df,
            remain_year_df=remain_year_df,
            balance_df=balance_df,
            stock_value_df=stock_value_df,
            sector_label=sector_label,
            dates=dates,
        )
    )

    out_path = Path(config.output_xlsx).resolve()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        info = pd.DataFrame(
            [
                {"key": "数据源", "value": f"parquet: {Path(config.parquet_root).resolve()}"},
                {"key": "分组口径", "value": "板块/股债型/平价/债项评级/剩余期限/余额/正股市值"},
                {"key": "分箱规则", "value": "左闭右开"},
                {"key": "过滤条件", "value": "不加过滤"},
                {"key": "统计量", "value": "中位数（median）"},
                {"key": "正股波动率字段", "value": config.stock_vol_sheet},
                {"key": "隐含波动率字段", "value": config.iv_sheet},
                {"key": "日期范围", "value": f"{dates[0].date()} ~ {dates[-1].date()} (n={len(dates)})"},
            ]
        )
        info.to_excel(writer, sheet_name="说明", index=False)

        for name, df in tables.items():
            sheet = name
            if len(sheet) > 31:
                sheet = sheet[:31]
            df.to_excel(writer, sheet_name=sheet, index=True)

    return out_path


if __name__ == "__main__":
    out = main(RunConfig())
    print(f"[ok] wrote: {out}")

