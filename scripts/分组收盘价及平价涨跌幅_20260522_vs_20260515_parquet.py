from __future__ import annotations

import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import numpy as np
import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from 底稿更新 import load_original_data


@dataclass(frozen=True)
class RunConfig:
    parquet_root: str = "data/转债个券历史序列"
    t0: str = "2026-05-15"
    t1: str = "2026-05-22"
    output_xlsx: str = "分组收盘价及平价涨跌幅_20260522_vs_20260515_parquet.xlsx"


def _require_col(df: pd.DataFrame, ts: pd.Timestamp, sheet: str) -> None:
    if ts not in df.columns:
        dates = [c for c in df.columns if isinstance(c, pd.Timestamp)]
        dates = sorted(dates)
        raise KeyError(f"[{sheet}] 未找到列 {ts.date()}，可用日期列示例: {dates[:5]} ... {dates[-5:]}")


def _as_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def _parity_ret(parity0: pd.Series, parity1: pd.Series) -> pd.Series:
    p0 = _as_num(parity0)
    p1 = _as_num(parity1)
    ret = p1 / p0 - 1.0
    ret[(~np.isfinite(ret)) | (p0 <= 0)] = np.nan
    return ret


def _mean_or_nan(x: pd.Series) -> float:
    x = pd.to_numeric(x, errors="coerce")
    v = x.dropna()
    if v.empty:
        return float("nan")
    return float(v.mean())


def _count_nonnull(*cols: pd.Series) -> int:
    if not cols:
        return 0
    ok = pd.Series(True, index=cols[0].index)
    for c in cols:
        ok &= pd.notna(c)
    return int(ok.sum())


def _build_group_table(
    group_label: pd.Series,
    close0: pd.Series,
    close1: pd.Series,
    parity0: pd.Series,
    parity1: pd.Series,
    label_order: Optional[List[str]] = None,
) -> pd.DataFrame:
    df = pd.DataFrame(
        {
            "group": group_label.astype("object"),
            "close0": _as_num(close0),
            "close1": _as_num(close1),
            "parity0": _as_num(parity0),
            "parity1": _as_num(parity1),
        }
    )
    df["parity_ret"] = _parity_ret(df["parity0"], df["parity1"])

    df = df.dropna(subset=["group"])

    rows: List[dict] = []
    for g, sub in df.groupby("group", sort=False):
        close_sub = sub.dropna(subset=["close0", "close1"])
        parity_sub = sub.dropna(subset=["parity_ret"])
        close0_mean = _mean_or_nan(close_sub["close0"])
        close1_mean = _mean_or_nan(close_sub["close1"])
        close_ret = float("nan")
        if np.isfinite(close0_mean) and np.isfinite(close1_mean) and close0_mean != 0.0:
            close_ret = close1_mean / close0_mean - 1.0
        rows.append(
            {
                "分组": str(g),
                "N(分组基准日有值)": int(len(sub)),
                "N(收盘价双日有效)": int(len(close_sub)),
                "收盘价(2026-05-15)": close0_mean,
                "收盘价(2026-05-22)": close1_mean,
                "收盘价涨跌幅": close_ret,
                "N(平价涨跌幅有效)": int(len(parity_sub)),
                "平价涨跌幅(均值,先个券后分组)": _mean_or_nan(parity_sub["parity_ret"]),
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        out = pd.DataFrame(
            columns=[
                "分组",
                "N(分组基准日有值)",
                "N(收盘价双日有效)",
                "收盘价(2026-05-15)",
                "收盘价(2026-05-22)",
                "收盘价涨跌幅",
                "N(平价涨跌幅有效)",
                "平价涨跌幅(均值,先个券后分组)",
            ]
        )
        return out
    if label_order:
        out["__ord__"] = pd.Categorical(out["分组"], categories=label_order, ordered=True)
        out = out.sort_values("__ord__", na_position="last").drop(columns="__ord__")
    else:
        out = out.sort_values("分组")
    return out.reset_index(drop=True)


def _bucket_left_closed_right_open(x: pd.Series, bins: List[tuple], labels: List[str]) -> pd.Series:
    xv = _as_num(x)
    out = pd.Series(np.nan, index=x.index, dtype="object")
    for (lo, hi), lb in zip(bins, labels):
        out[(xv >= lo) & (xv < hi)] = lb
    return out


def main(config: RunConfig) -> Path:
    t0 = pd.Timestamp(config.t0)
    t1 = pd.Timestamp(config.t1)

    data = load_original_data(source_type="parquet", parquet_root=config.parquet_root, force_refresh=False)

    required_sheets = [
        "收盘价",
        "平价",
        "申万行业",  # 来自 总表
        "平价底价溢价率",
        "债项评级",
        "剩余期限",
        "余额",
        "正股市值",
    ]

    if "总表" not in data:
        raise KeyError("未在parquet数据中找到 `总表`（用于读取申万行业）。")

    close_df = data["收盘价"]
    parity_df = data["平价"]
    stock_bond_df = data["平价底价溢价率"]
    rating_df = data["债项评级"]
    remain_year_df = data["剩余期限"]
    balance_df = data["余额"]
    stock_value_df = data["正股市值"]

    for sheet, df in [
        ("收盘价", close_df),
        ("平价", parity_df),
        ("平价底价溢价率", stock_bond_df),
        ("债项评级", rating_df),
        ("剩余期限", remain_year_df),
        ("余额", balance_df),
        ("正股市值", stock_value_df),
    ]:
        _require_col(df, t0, sheet)
        _require_col(df, t1, sheet)

    close0 = close_df[t0]
    close1 = close_df[t1]
    parity0 = parity_df[t0]
    parity1 = parity_df[t1]

    # ====== 分组：板块（申万行业 -> 5 大类） ======
    total_df = data["总表"].copy()
    if "__row_id" in total_df.columns:
        total_df = total_df.set_index("__row_id")
    industry_s = total_df.get("申万行业")
    if industry_s is None:
        raise KeyError("`总表` 中未找到 `申万行业` 字段。")
    industry_s = industry_s.reindex(close_df.index)

    sector_map: Dict[str, List[str]] = {
        "科技": ["传媒", "电子", "国防军工", "计算机", "通信"],
        "金融": ["非银金融", "银行"],
        "制造": ["电力设备", "机械设备", "汽车", "轻工制造"],
        "消费": ["农林牧渔", "纺织服饰", "家用电器", "商贸零售", "社会服务", "食品饮料", "医药生物", "美容护理"],
        "周期": ["基础化工", "钢铁", "公用事业", "环保", "建筑材料", "建筑装饰", "交通运输", "煤炭", "石油石化", "有色金属"],
    }
    sector_label = pd.Series(np.nan, index=industry_s.index, dtype="object")
    for sec, inds in sector_map.items():
        sector_label[industry_s.isin(inds)] = sec

    # ====== 分组：股债型（按平价底价溢价率阈值） ======
    sb0 = _as_num(stock_bond_df[t0])
    sb_label = pd.Series(np.nan, index=sb0.index, dtype="object")
    sb_label[sb0 < -20] = "偏债型"
    sb_label[(sb0 >= -20) & (sb0 < 20)] = "平衡型"
    sb_label[sb0 >= 20] = "偏股型"

    # ====== 分组：平价（左闭右开） ======
    plain0 = _as_num(parity0)
    plain_bins = [(70, 90), (90, 110), (110, 130), (130, 150)]
    plain_labels = ["70-90", "90-110", "110-130", "130-150"]
    plain_label = _bucket_left_closed_right_open(plain0, plain_bins, plain_labels)

    # ====== 分组：债项评级 ======
    rating0_raw = rating_df[t0].astype("object")
    rating1_raw = rating_df[t1].astype("object")
    rating_for_group = rating0_raw.copy()
    invalid0 = rating_for_group.isna() | rating_for_group.astype(str).str.strip().isin({"0", "0.0", ""})
    rating_for_group[invalid0] = rating1_raw[invalid0]

    rating_label = pd.Series(np.nan, index=rating_for_group.index, dtype="object")
    rating_label[rating_for_group.isin(["AAA", "AA+"])] = "AAA/AA+"
    rating_label[rating_for_group.isin(["AA", "AA-"])] = "AA/AA-"
    rating_label[rating_for_group.isin(["A+", "A"])] = "A/A-"

    # ====== 分组：剩余期限（左闭右开） ======
    ry0 = _as_num(remain_year_df[t0])
    ry_bins = [(0, 1), (1, 2), (2, 3), (3, 4), (4, 5), (5, 6)]
    ry_labels = ["0-1", "1-2", "2-3", "3-4", "4-5", "5-6"]
    ry_label = _bucket_left_closed_right_open(ry0, ry_bins, ry_labels)

    # ====== 分组：余额（左闭右开） ======
    bal0 = _as_num(balance_df[t0])
    bal_bins = [(0, 3), (3, 10), (10, 20), (20, 50)]
    bal_labels = ["0-3", "3-10", "10-20", "20-50"]
    bal_label = _bucket_left_closed_right_open(bal0, bal_bins, bal_labels)
    bal_label[bal0 >= 50] = "50+"

    # ====== 分组：正股市值（左闭右开） ======
    sv0 = _as_num(stock_value_df[t0])
    sv_bins = [(0, 50), (50, 300)]
    sv_labels = ["0-50", "50-300"]
    sv_label = _bucket_left_closed_right_open(sv0, sv_bins, sv_labels)
    sv_label[sv0 >= 300] = "300+"

    tables = {
        "板块": _build_group_table(sector_label, close0, close1, parity0, parity1, label_order=["科技", "金融", "制造", "消费", "周期"]),
        "股债型": _build_group_table(sb_label, close0, close1, parity0, parity1, label_order=["偏债型", "平衡型", "偏股型"]),
        "平价": _build_group_table(plain_label, close0, close1, parity0, parity1, label_order=plain_labels),
        "债项评级": _build_group_table(rating_label, close0, close1, parity0, parity1, label_order=["AAA/AA+", "AA/AA-", "A/A-"]),
        "剩余期限": _build_group_table(ry_label, close0, close1, parity0, parity1, label_order=ry_labels),
        "余额": _build_group_table(bal_label, close0, close1, parity0, parity1, label_order=["0-3", "3-10", "10-20", "20-50", "50+"]),
        "正股市值": _build_group_table(sv_label, close0, close1, parity0, parity1, label_order=["0-50", "50-300", "300+"]),
    }

    out_path = Path(config.output_xlsx).resolve()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        info = pd.DataFrame(
            [
                {"key": "数据源", "value": f"parquet: {Path(config.parquet_root).resolve()}"},
                {"key": "基准日(t0)", "value": str(t0.date())},
                {"key": "对比日(t1)", "value": str(t1.date())},
                {"key": "平价涨跌幅口径", "value": "先算个券涨跌幅，再分组取均值"},
                {"key": "分箱口径", "value": "左闭右开"},
                {"key": "过滤条件", "value": "不加过滤"},
            ]
        )
        info.to_excel(writer, sheet_name="说明", index=False)
        for sheet, df in tables.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

    return out_path


if __name__ == "__main__":
    out = main(RunConfig())
    print(f"[ok] wrote: {out}")
