# -*- coding: utf-8 -*-
"""日报/周报 iFinD 数据拉取与计算脚本。

数据流原则（相对 notebook）：
- API 拉取 → 内存宽表 ``raw_data`` → 清洗/汇总/拟合/JS/文本 全程在内存中传递；
- ``MMDD数据更新.xlsx`` 仅作落盘备份，**同一次运行内不得** ``read_excel`` 再读回做计算；
- ``清理后`` / ``清理后剔妖`` 统计文件只含汇总 sheet（从全样本余额加权起），不含余额/收盘价等底稿。
- A 轨清洗：上市日前、最后交易日后置空，并按收盘价缺失同步全表（与 notebook 一致，不用换手率筛选）。

运行：python 【更新】日报数据更新.py
"""

from __future__ import annotations

import os
import shutil
import subprocess
import time as _time
from collections import Counter
from configparser import ConfigParser
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Literal, Optional

import matplotlib
matplotlib.use("Agg")
from matplotlib import font_manager
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from iFinDPy import *
from scipy.integrate import quad
from scipy.optimize import curve_fit
from tqdm import tqdm

# 内存 sheet 字典：中文名 -> 宽表 DataFrame
SheetDict = dict[str, pd.DataFrame]

_WORKSPACE_ROOT = Path(__file__).resolve().parents[2]
_SCRIPT_DIR = str(_WORKSPACE_ROOT)
RUNS_DAILY_ROOT = _WORKSPACE_ROOT / "runs" / "daily"
_KAITI_CANDIDATES = (
    "assets/fonts/KaiTi_GB2312.ttf",
    "KaiTi_GB2312.TTF",
    "KaiTi_GB2312.otf",
    "KaiTi_GB2312",
    "kaiti_gb2312.ttf",
)


def _configure_matplotlib_chinese_font() -> None:
    for name in _KAITI_CANDIDATES:
        path = os.path.join(_SCRIPT_DIR, name)
        if not os.path.isfile(path):
            continue
        font_manager.fontManager.addfont(path)
        family = font_manager.FontProperties(fname=path).get_name()
        plt.rcParams["font.sans-serif"] = [family, "KaiTi", "SimHei", "Microsoft YaHei"]
        plt.rcParams["axes.unicode_minus"] = False
        return
    plt.rcParams["font.sans-serif"] = ["KaiTi_GB2312", "KaiTi", "SimHei", "Microsoft YaHei"]
    plt.rcParams["axes.unicode_minus"] = False


_configure_matplotlib_chinese_font()

# ================== 统计清洗与汇总（内存） ==================

WIDE_META_COLS = ("代码", "名称", "上市日期", "评级")
STATIC_SHEETS = frozenset({"总表", "发行规模"})
SKIP_MASK_SHEETS = STATIC_SHEETS

TRACK_B_PREMIUM_MIN = 50.0
TRACK_B_CLOSE_MIN = 150.0

# 百元拟合样本筛选（notebook cell 180 / 182，与统计 A 轨「不用换手率」无关）
FIT_PLAIN_MIN = 70.0
FIT_PLAIN_MAX = 130.0
FIT_TURNOVER_MAX = 50.0

CLOSE_QUANTILE_PER = [0.05, 0.25, 0.5, 0.75, 0.8, 0.9]
YTM_QUANTILE_PER = CLOSE_QUANTILE_PER

SheetDict = dict[str, pd.DataFrame]


def date_cols(df: pd.DataFrame) -> list[str]:
    return [str(c) for c in df.columns if c not in WIDE_META_COLS]


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def balance_weighted_mean(values: pd.Series, weights: pd.Series) -> float:
    v, w = _num(values), _num(weights)
    m = v.notna() & w.notna() & (w > 0)
    if not m.any():
        return np.nan
    return float((v[m] * w[m]).sum() / w[m].sum())


def avg_classify(pairs: Iterable[tuple[float, float]]) -> float:
    """pairs: (metric, balance)"""
    rows = [(a, b) for a, b in pairs if not (pd.isna(a) or pd.isna(b))]
    if not rows:
        return np.nan
    tm = pd.DataFrame(rows, columns=["m", "w"])
    wsum = tm["w"].sum()
    if wsum == 0 or pd.isna(wsum):
        return np.nan
    return float((tm["m"] / wsum * tm["w"]).sum())


def copy_sheet_dict(data: SheetDict) -> SheetDict:
    return {k: v.copy() for k, v in data.items()}


def apply_listing_date_mask(data: SheetDict) -> None:
    """上市日以前各指标置 NaN（与 notebook 一致）。"""
    for name, df in data.items():
        if name in SKIP_MASK_SHEETS:
            continue
        ipo = pd.to_datetime(df["上市日期"], errors="coerce")
        for d in date_cols(df):
            col_date = pd.to_datetime(d, errors="coerce")
            if pd.isna(col_date):
                continue
            bad = ipo.notna() & (col_date < ipo)
            df.loc[bad.values, d] = np.nan


def apply_delist_close_mask(data: SheetDict) -> None:
    """最后交易日后置 NaN，并按当日收盘价缺失同步全表（与 notebook 一致）。"""
    zb = data["总表"].set_index("代码")
    last_trade = pd.to_datetime(zb["最后交易日"], errors="coerce")
    close = data["收盘价"]
    codes = close["代码"].astype(str)
    ltd = last_trade.reindex(codes)

    for d in date_cols(close):
        col_date = pd.to_datetime(d, errors="coerce")
        if pd.isna(col_date):
            continue
        bad_delist = ltd.notna() & (col_date > ltd)
        close.loc[bad_delist.values, d] = np.nan

    for d in date_cols(close):
        bad = _num(close[d]).isna()
        for name, df in data.items():
            if name in SKIP_MASK_SHEETS or d not in df.columns:
                continue
            df.loc[bad.values, d] = np.nan


def apply_listing_delist_masks(data: SheetDict) -> None:
    """A 轨：上市日 + 退市日（收盘价缺失）清洗。"""
    apply_listing_date_mask(data)
    apply_delist_close_mask(data)


def apply_fit_plain_range_mask(data: SheetDict) -> None:
    """拟合前：平价不在 [70, 130] 的券当日全表置 NaN（notebook「70-130平价样本筛选」）。"""
    plain = data["平价"]
    for d in date_cols(plain):
        p = _num(plain[d])
        bad = p.notna() & ((p < FIT_PLAIN_MIN) | (p > FIT_PLAIN_MAX))
        if not bad.any():
            continue
        for name, df in data.items():
            if name in SKIP_MASK_SHEETS or d not in df.columns:
                continue
            df.loc[bad.values, d] = np.nan


def apply_fit_high_turnover_mask(data: SheetDict) -> None:
    """拟合前：换手率>50 的券当日全表置 NaN（notebook「换手率剔除」）。"""
    if "换手率" not in data:
        return
    turn = data["换手率"]
    for d in date_cols(turn):
        bad = _num(turn[d]) > FIT_TURNOVER_MAX
        if not bad.any():
            continue
        for name, df in data.items():
            if name in SKIP_MASK_SHEETS or d not in df.columns:
                continue
            df.loc[bad.values, d] = np.nan


def apply_yao_mask(data: SheetDict) -> None:
    """B 轨：转股溢价率>50 且 收盘价>150。"""
    dates = date_cols(data["收盘价"])
    for d in dates:
        bad = (_num(data["转股溢价率"][d]) > TRACK_B_PREMIUM_MIN) & (
            _num(data["收盘价"][d]) > TRACK_B_CLOSE_MIN
        )
        for name, df in data.items():
            if name in SKIP_MASK_SHEETS or d not in df.columns:
                continue
            df.loc[bad.values, d] = np.nan


def prepare_masked_data(data: SheetDict, *, apply_track_b: bool) -> SheetDict:
    out = copy_sheet_dict(data)
    apply_listing_delist_masks(out)
    if apply_track_b:
        apply_yao_mask(out)
    return out


def stat_full_sample_balance(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    need = [
        "收盘价",
        "平价",
        "转股溢价率",
        "纯债溢价率",
        "纯债价值",
        "隐含波动率",
        "YTM",
        "次新券相对上市首日涨跌幅",
        "次新券转股溢价率",
    ]
    tm = data["余额"]
    lines = []
    for c in colum:
        line = []
        for name in need:
            if name == "隐含波动率":
                tmp2 = _num(data[name][c])
                tmp1 = _num(tm[c])
                tmp1 = tmp1.copy()
                tmp1[tmp2 == 0] = 0
                try:
                    line.append(float((tmp2 * tmp1 / tmp1.sum()).sum()))
                except Exception:
                    line.append(np.nan)
            elif name in ("次新券相对上市首日涨跌幅", "次新券转股溢价率"):
                # 仅次新券行有值：用当日余额对次新指标加权（列名虽带「余额加权」，不能用全样本算术均值）
                line.append(
                    balance_weighted_mean(
                        _num(data[name][c]).replace(0, np.nan),
                        tm[c],
                    )
                )
            else:
                line.append(
                    balance_weighted_mean(data[name][c], tm[c])
                )
        lines.append(line)
    cols = [f"{n}余额加权" for n in need]
    df = pd.DataFrame(lines, columns=cols, index=colum)
    order = [
        "转股溢价率余额加权",
        "收盘价余额加权",
        "平价余额加权",
        "纯债价值余额加权",
        "纯债溢价率余额加权",
        "隐含波动率余额加权",
        "YTM余额加权",
        "次新券相对上市首日涨跌幅余额加权",
        "次新券转股溢价率余额加权",
    ]
    return df[order]


def stat_full_sample_mean(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    need = [
        "收盘价",
        "平价",
        "转股溢价率",
        "纯债溢价率",
        "纯债价值",
        "隐含波动率",
        "YTM",
        "次新券相对上市首日涨跌幅",
        "次新券转股溢价率",
    ]
    lines = []
    for c in colum:
        line = []
        for name in need:
            line.append(float(_num(data[name][c]).mean()))
        lines.append(line)
    cols = [f"{n}算数均值" for n in need]
    df = pd.DataFrame(lines, columns=cols, index=colum)
    order = [
        "转股溢价率算数均值",
        "收盘价算数均值",
        "平价算数均值",
        "纯债价值算数均值",
        "纯债溢价率算数均值",
        "隐含波动率算数均值",
        "YTM算数均值",
        "次新券相对上市首日涨跌幅算数均值",
        "次新券转股溢价率算数均值",
    ]
    return df[order]


def stat_rating_premium(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm1, tm2, tm3 = data["债项评级"], data["转股溢价率"], data["余额"]
    lines = []
    for c in colum:
        buckets = {k: [] for k in ["AAA", "AA+", "AA", "AA-", "A+", "A及以下"]}
        for a, b, bal in zip(tm1[c], tm2[c], tm3[c]):
            if pd.isna(a) or pd.isna(b) or pd.isna(bal):
                continue
            key = str(a).strip()
            if not key or key.lower() == "nan":
                continue
            if key not in buckets:
                key = "A及以下"
            buckets[key].append((float(b), float(bal)))
        lines.append(
            tuple(avg_classify(buckets[k]) for k in ["AAA", "AA+", "AA", "AA-", "A+", "A及以下"])
        )
    return pd.DataFrame(
        lines,
        columns=["AAA", "AA+", "AA", "AA-", "A+", "A及以下"],
        index=colum,
    )


def stat_size_premium(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    """按当日余额动态分组，并计算各组转股溢价率的余额加权均值。"""
    tm1, tm2 = data["余额"], data["转股溢价率"]
    order = ["50亿以上", "20-50亿（含50亿）", "10-20亿（含20亿）", "3-10亿（含10亿）", "3亿以下"]
    lines = []
    for c in colum:
        balance = _num(tm1[c])
        row = []
        for lb in order:
            if lb == "3亿以下":
                idx = balance <= 3
            elif lb == "3-10亿（含10亿）":
                idx = (balance > 3) & (balance <= 10)
            elif lb == "10-20亿（含20亿）":
                idx = (balance > 10) & (balance <= 20)
            elif lb == "20-50亿（含50亿）":
                idx = (balance > 20) & (balance <= 50)
            else:
                idx = balance > 50
            row.append(balance_weighted_mean(tm2.loc[idx, c], balance.loc[idx]))
        lines.append(row)
    return pd.DataFrame(lines, columns=order, index=colum)


def stat_mv_bucket(
    data: SheetDict,
    colum: list[str],
    metric: str,
    col_names: list[str],
) -> pd.DataFrame:
    tm1, tm2, tm3 = data["正股市值"], data[metric], data["余额"]

    def bucket(a):
        if pd.isna(a):
            return None
        a = float(a)
        if a <= 30:
            return col_names[0]
        if a <= 50:
            return col_names[1]
        if a <= 100:
            return col_names[2]
        if a <= 500:
            return col_names[3]
        if a > 500:
            return col_names[4]
        return None

    lines = []
    for c in colum:
        buckets = {k: [] for k in col_names}
        for a, b, bal in zip(tm1[c], tm2[c], tm3[c]):
            if pd.isna(a) or pd.isna(b) or pd.isna(bal):
                continue
            bk = bucket(a)
            if bk:
                buckets[bk].append((float(b), float(bal)))
        lines.append(tuple(avg_classify(buckets[k]) for k in col_names))
    return pd.DataFrame(lines, columns=col_names, index=colum)


def stat_parity_premium(data: SheetDict, colum: list[str], metric: str) -> pd.DataFrame:
    tm1, tm2, tm3 = data["平价"], data[metric], data["余额"]
    cols = ["80", "90", "100", "110", "120", "130", "inf"]
    lines = []
    for c in colum:
        buckets = {k: [] for k in cols}

        def put(a, b, bal):
            if pd.isna(a) or pd.isna(b) or pd.isna(bal):
                return
            a = float(a)
            if a <= 80:
                buckets["80"].append((b, bal))
            elif a <= 90:
                buckets["90"].append((b, bal))
            elif a <= 100:
                buckets["100"].append((b, bal))
            elif a <= 110:
                buckets["110"].append((b, bal))
            elif a <= 120:
                buckets["120"].append((b, bal))
            elif a <= 130:
                buckets["130"].append((b, bal))
            else:
                buckets["inf"].append((b, bal))

        for a, b, bal in zip(tm1[c], tm2[c], tm3[c]):
            put(a, b, bal)
        lines.append(tuple(avg_classify(buckets[k]) for k in cols))
    df = pd.DataFrame(lines, columns=cols, index=colum)
    return df[["inf", "130", "120", "110", "100", "90", "80"]]


def stat_close_quantile(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm = data["收盘价"]
    lines = []
    for c in colum:
        nums = [float(x) for x in _num(tm[c]) if not pd.isna(x)]
        if nums:
            q = pd.Series(nums).quantile(q=CLOSE_QUANTILE_PER).tolist()
            lines.append(q + [float(np.mean(nums))])
        else:
            lines.append([np.nan] * (len(CLOSE_QUANTILE_PER) + 1))
    return pd.DataFrame(lines, columns=CLOSE_QUANTILE_PER + ["均值"], index=colum)


def stat_ytm_quantile(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm = data["YTM"]
    lines = []
    for i, c in enumerate(colum):
        nums = [float(x) for x in _num(tm[c]) if not pd.isna(x)]
        if c == colum[-1] and len(colum) > 1 and nums:
            prev = colum[-2]
            if prev in tm.columns:
                cur = _num(tm[c])
                prv = _num(tm[prev])
                if cur.equals(prv):
                    nums = []
        if nums:
            q = pd.Series(nums).quantile(q=YTM_QUANTILE_PER).tolist()
            lines.append(q + [float(np.mean(nums))])
        else:
            lines.append([np.nan] * (len(YTM_QUANTILE_PER) + 1))
    return pd.DataFrame(lines, columns=YTM_QUANTILE_PER + ["均值"], index=colum)


def stat_pure_bond_analysis(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    bal, pure, close = data["余额"], data["纯债价值"], data["收盘价"]
    lines = []
    for c in colum:
        b, p, cl = _num(bal[c]), _num(pure[c]), _num(close[c])
        m = p.notna() & cl.notna() & b.notna()
        bw = balance_weighted_mean(p[m], b[m])
        mask = m & (p > cl)
        cnt = int(mask.sum())
        total = int(cl.notna().sum())
        ratio = cnt / total if total else np.nan
        lines.append((ratio, bw, cnt, total))
    return pd.DataFrame(
        lines,
        columns=["纯债>转债价格比例", "纯债余额加权", "纯债>转债价格数量", "转债总数"],
        index=colum,
    )


def stat_floor_balance_weighted(
    data: SheetDict, colum: list[str], metric: str
) -> pd.DataFrame:
    tm1, tm2, tm3 = data["平价底价溢价率"], data[metric], data["余额"]
    cols = ["偏股", "偏债", "平衡"]
    lines = []
    for c in colum:
        buckets = {k: [] for k in cols}
        for a, b, bal in zip(tm1[c], tm2[c], tm3[c]):
            if pd.isna(a) or pd.isna(b) or pd.isna(bal):
                continue
            a = float(a)
            if a > 20:
                buckets["偏股"].append((b, bal))
            elif a < -20:
                buckets["偏债"].append((b, bal))
            else:
                buckets["平衡"].append((b, bal))
        lines.append(tuple(avg_classify(buckets[k]) for k in cols))
    return pd.DataFrame(lines, columns=cols, index=colum)


def stat_price_premium_quadrant_premium(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm_a, tm_b, tm_w = data["转股溢价率"], data["收盘价"], data["余额"]
    cols = ["双高", "低价高估", "双低", "高价低估"]
    lines = []
    for c in colum:
        buckets = {k: [] for k in cols}
        for a, b, bal in zip(tm_a[c], tm_b[c], tm_w[c]):
            if pd.isna(a) or pd.isna(b) or pd.isna(bal):
                continue
            a, b = float(a), float(b)
            if a > 20 and b > 135:
                buckets["双高"].append((a, bal))
            elif a > 20 and b <= 135:
                buckets["低价高估"].append((a, bal))
            elif a <= 20 and b <= 135:
                buckets["双低"].append((a, bal))
            else:
                buckets["高价低估"].append((a, bal))
        lines.append(tuple(avg_classify(buckets[k]) for k in cols))
    df = pd.DataFrame(lines, columns=cols, index=colum)
    return df[["双高", "双低", "低价高估", "高价低估"]]


def stat_break_floor(data: SheetDict, colum: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    price, pure, bal = data["收盘价"], data["纯债价值"], data["余额"]
    r1 = pd.DataFrame(index=colum, columns=["破底", "未破底"])
    r2 = pd.DataFrame(index=colum, columns=["破底余额总和", "未破底余额总和"])
    for c in colum:
        cl, pu, b = _num(price[c]), _num(pure[c]), _num(bal[c])
        m = cl.notna() & pu.notna()
        mask = m & (cl < pu)
        r1.loc[c, "破底"] = int(mask.sum())
        r1.loc[c, "未破底"] = int((m & ~mask).sum())
        r2.loc[c, "破底余额总和"] = float(b[mask].sum())
        r2.loc[c, "未破底余额总和"] = float(b[m & ~mask].sum())
    r1["破底占比"] = r1["破底"] / (r1["破底"] + r1["未破底"])
    r1["未破底占比"] = 1 - r1["破底占比"]
    r2["破底余额占比"] = r2["破底余额总和"] / (
        r2["破底余额总和"] + r2["未破底余额总和"]
    )
    r2["未破底余额占比"] = 1 - r2["破底余额占比"]
    return r1, r2


def stat_daily_count_balance(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    """清洗后：收盘价有效且余额>0 的券（不再按换手率筛选）。"""
    bal, close = data["余额"], data["收盘价"]
    rows = []
    for c in colum:
        b, cl = _num(bal[c]), _num(close[c])
        valid = cl.notna() & b.notna() & (b > 0)
        rows.append({"转债个数": int(valid.sum()), "总余额": float(b[valid].sum())})
    return pd.DataFrame(rows, index=colum)


def stat_close_bucket_premium(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm1, tm2, tm3 = data["收盘价"], data["转股溢价率"], data["余额"]
    cols = ["90以下", "90-100", "100-110", "110-120", "120-130", "130-150", "150以上"]
    lines = []
    for c in colum:
        buckets = {k: [] for k in cols}

        def put(a, b, bal):
            if pd.isna(a) or pd.isna(b) or pd.isna(bal):
                return
            a = float(a)
            if a <= 90:
                buckets["90以下"].append((b, bal))
            elif a <= 100:
                buckets["90-100"].append((b, bal))
            elif a <= 110:
                buckets["100-110"].append((b, bal))
            elif a <= 120:
                buckets["110-120"].append((b, bal))
            elif a <= 130:
                buckets["120-130"].append((b, bal))
            elif a <= 150:
                buckets["130-150"].append((b, bal))
            else:
                buckets["150以上"].append((b, bal))

        for a, b, bal in zip(tm1[c], tm2[c], tm3[c]):
            put(a, b, bal)
        lines.append(tuple(avg_classify(buckets[k]) for k in cols))
    return pd.DataFrame(lines, columns=cols, index=colum)


def stat_price_bucket_ytm(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm1, tm2, tm3 = data["收盘价"], data["YTM"], data["余额"]
    cols = ["90", "100", "110", "120", "130", "150", "inf"]
    lines = []
    for c in colum:
        buckets = {k: [] for k in cols}
        for a, b, bal in zip(tm1[c], tm2[c], tm3[c]):
            if pd.isna(a) or pd.isna(b) or pd.isna(bal):
                continue
            a = float(a)
            if a <= 90:
                buckets["90"].append((b, bal))
            elif a <= 100:
                buckets["100"].append((b, bal))
            elif a <= 110:
                buckets["110"].append((b, bal))
            elif a <= 120:
                buckets["120"].append((b, bal))
            elif a <= 130:
                buckets["130"].append((b, bal))
            elif a <= 150:
                buckets["150"].append((b, bal))
            else:
                buckets["inf"].append((b, bal))
        lines.append(tuple(avg_classify(buckets[k]) for k in cols))
    return pd.DataFrame(lines, columns=cols, index=colum)


def stat_aaa_ytm_median(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm1, tm2, tm3 = data["债项评级"], data["YTM"], data["剩余期限"]
    lines = []
    for c in colum:
        k1, k2 = [], []
        for a, b, rem in zip(tm1[c], tm2[c], tm3[c]):
            if pd.isna(a) or pd.isna(b) or pd.isna(rem):
                continue
            if str(a) == "AAA" and float(rem) >= 1:
                k1.append(float(b))
            else:
                k2.append(float(b))
        t1 = float(pd.Series(k1).median()) if k1 else np.nan
        t2 = float(pd.Series(k2).median()) if k2 else np.nan
        lines.append((t1, t2))
    return pd.DataFrame(lines, columns=["AAA一年内", "其他"], index=colum)


def stat_parity_interval_ratio(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm_label = [
        "小于50（包含50）",
        "50-70（包含70）",
        "70-80（包含80）",
        "80-90（包含90）",
        "90-100（包含100）",
        "大于100",
    ]

    def get_label(i: float) -> str:
        if i <= 50:
            return tm_label[0]
        if i <= 70:
            return tm_label[1]
        if i <= 80:
            return tm_label[2]
        if i <= 90:
            return tm_label[3]
        if i <= 100:
            return tm_label[4]
        return tm_label[5]

    tm = data["平价"]
    dfs = []
    for c in colum:
        labels = [get_label(float(j)) for j in _num(tm[c]) if not pd.isna(j)]
        counter = Counter(labels)
        row = {lb: counter.get(lb, 0) for lb in tm_label}
        total = sum(row.values())
        df_row = dict(row)
        for lb in tm_label:
            df_row[f"{lb}_比例"] = row[lb] / total if total else np.nan
        dfs.append(pd.DataFrame([df_row], index=[c]))
    out = pd.concat(dfs)
    cols = tm_label + [f"{lb}_比例" for lb in tm_label]
    return out[cols]


def stat_close_interval_ratio(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm_label = [
        "小于80（包含80）",
        "80-90（包含90）",
        "90-100（包含100）",
        "100-110（包含110）",
        "110-120（包含120）",
        "120-130（包含130）",
        "130-150（包含150）",
        "大于150",
    ]

    def get_label(i: float) -> str:
        if i <= 80:
            return tm_label[0]
        if i <= 90:
            return tm_label[1]
        if i <= 100:
            return tm_label[2]
        if i <= 110:
            return tm_label[3]
        if i <= 120:
            return tm_label[4]
        if i <= 130:
            return tm_label[5]
        if i <= 150:
            return tm_label[6]
        return tm_label[7]

    tm = data["收盘价"]
    dfs = []
    for c in colum:
        labels = [get_label(float(j)) for j in _num(tm[c]) if not pd.isna(j)]
        counter = Counter(labels)
        row = {lb: counter.get(lb, 0) for lb in tm_label}
        total = sum(row.values())
        df_row = dict(row)
        for lb in tm_label:
            df_row[f"{lb}_比例"] = row[lb] / total if total else np.nan
        dfs.append(pd.DataFrame([df_row], index=[c]))
    out = pd.concat(dfs)
    cols = tm_label + [f"{lb}_比例" for lb in tm_label]
    return out[cols]


def build_all_statistics(data: SheetDict) -> SheetDict:
    """生成并集内全部汇总 sheet（不含底稿宽表）。"""
    colum = date_cols(data["余额"])
    stats: SheetDict = {}

    stats["全样本余额"] = stat_full_sample_balance(data, colum)
    stats["全样本算数"] = stat_full_sample_mean(data, colum)
    stats["评级分类余额加权转股溢价率"] = stat_rating_premium(data, colum)
    stats["规模分类余额加权转股溢价率"] = stat_size_premium(data, colum)
    stats["正股市值分类转股溢价率"] = stat_mv_bucket(
        data,
        colum,
        "转股溢价率",
        ["30亿以下(含)", "30-50亿（含）", "50-100亿（含）", "100-500亿（含）", "500亿以上"],
    )
    stats["正股市值分类收盘价"] = stat_mv_bucket(
        data,
        colum,
        "收盘价",
        ["30亿以下(含)", "30-50亿（含）", "50-100亿（含）", "100-500亿（含）", "500亿以上"],
    )
    stats["平价分类余额加权转股溢价率"] = stat_parity_premium(
        data, colum, "转股溢价率"
    )
    stats["平价分类余额加权收盘价"] = stat_parity_premium(data, colum, "收盘价")
    stats["平价区间数量比例"] = stat_parity_interval_ratio(data, colum)
    stats["收盘价区间数量比例"] = stat_close_interval_ratio(data, colum)
    stats["收盘价分位数统计"] = stat_close_quantile(data, colum)
    stats["YTM分位数统计"] = stat_ytm_quantile(data, colum)
    stats["纯债分析"] = stat_pure_bond_analysis(data, colum)
    stats["平底分类余额加权收盘价"] = stat_floor_balance_weighted(
        data, colum, "收盘价"
    )
    stats["平底分类余额加权转股溢价率"] = stat_floor_balance_weighted(
        data, colum, "转股溢价率"
    )
    stats["价格溢价率分类余额加权收盘价"] = _stat_quadrant_close(data, colum)
    stats["价格溢价率分类余额加权溢价率"] = stat_price_premium_quadrant_premium(
        data, colum
    )
    r1, r2 = stat_break_floor(data, colum)
    stats["每日破底统计"] = r1
    stats["每日破底余额统计"] = r2
    stats["每日数量总余额_换手非零"] = stat_daily_count_balance(data, colum)
    stats["收盘价分类余额加权转股溢价率"] = stat_close_bucket_premium(data, colum)
    stats["价格分类余额加权YTM"] = stat_price_bucket_ytm(data, colum)
    stats["AAA评级YTM中位数"] = stat_aaa_ytm_median(data, colum)
    for k in stats:
        stats[k] = stats[k].copy()
        stats[k].index.name = "日期"
    return stats


def _stat_quadrant_close(data: SheetDict, colum: list[str]) -> pd.DataFrame:
    tm_a, tm_b, tm_w = data["转股溢价率"], data["收盘价"], data["余额"]
    cols = ["双高", "低价高估", "双低", "高价低估"]
    lines = []
    for c in colum:
        buckets = {k: [] for k in cols}
        for a, b, bal in zip(tm_a[c], tm_b[c], tm_w[c]):
            if pd.isna(a) or pd.isna(b) or pd.isna(bal):
                continue
            a, b = float(a), float(b)
            if a > 20 and b > 135:
                buckets["双高"].append((b, bal))
            elif a > 20 and b <= 135:
                buckets["低价高估"].append((b, bal))
            elif a <= 20 and b <= 135:
                buckets["双低"].append((b, bal))
            else:
                buckets["高价低估"].append((b, bal))
        lines.append(tuple(avg_classify(buckets[k]) for k in cols))
    df = pd.DataFrame(lines, columns=cols, index=colum)
    return df[["双高", "双低", "低价高估", "高价低估"]]


def compute_statistics(data: SheetDict, *, apply_track_b: bool) -> SheetDict:
    masked = prepare_masked_data(data, apply_track_b=apply_track_b)
    return build_all_statistics(masked)


# ================== 百元平价拟合与 JS（内存） ==================

FloorMode = Optional[Literal["balance", "bond", "stock"]]
RemainMode = Optional[tuple[float, float]]


def inverse_cubic(x, a, b, c, d):
    return a / np.power(x, 3) + b / np.power(x, 2) + c / x + d


def delta_cal(x, a, b, c, d):
    return (-2 * a / np.power(x, 3) - b / np.power(x, 2) + d + 100) / 100


def gamma_cal(x, a, b, c, d):
    return (6 * a / np.power(x, 4) + 2 * b / np.power(x, 3)) / 100


def prepare_fit_data(data: SheetDict) -> SheetDict:
    """拟合样本：上市/退市 + 平价 70–130 + 换手率≤50（与日报 notebook 拟合段一致，不剔妖）。"""
    out = copy_sheet_dict(data)
    apply_listing_delist_masks(out)
    apply_fit_plain_range_mask(out)
    apply_fit_high_turnover_mask(out)
    return out


def _matrix(wide: pd.DataFrame) -> pd.DataFrame:
    return wide.set_index("代码")[date_cols(wide)].apply(pd.to_numeric, errors="coerce")


def _fit_one_date(
    plain: pd.Series,
    prem: pd.Series,
    floor: Optional[pd.Series] = None,
    floor_mode: FloorMode = None,
    remain: Optional[pd.Series] = None,
    remain_range: RemainMode = None,
    *,
    with_greeks: bool = False,
) -> list:
    df_date = pd.concat([plain, prem], axis=1)
    if floor is not None:
        df_date = pd.concat([df_date, floor], axis=1)
    if remain is not None:
        df_date = pd.concat([df_date, remain], axis=1)

    df_date = df_date.dropna(how="all")
    y = pd.to_numeric(df_date.iloc[:, 1], errors="coerce")
    low, up = y.quantile(0.03), y.quantile(0.97)
    df_date = df_date[(y > low) & (y < up)]

    if floor_mode == "balance" and df_date.shape[1] >= 3:
        f = pd.to_numeric(df_date.iloc[:, 2], errors="coerce")
        df_date = df_date[(f >= -20) & (f <= 20)]
    elif floor_mode == "bond" and df_date.shape[1] >= 3:
        f = pd.to_numeric(df_date.iloc[:, 2], errors="coerce")
        df_date = df_date[f < -20]
    elif floor_mode == "stock" and df_date.shape[1] >= 3:
        f = pd.to_numeric(df_date.iloc[:, 2], errors="coerce")
        df_date = df_date[f > 20]

    if remain_range and df_date.shape[1] >= (4 if floor is not None else 3):
        rcol = -1
        r = pd.to_numeric(df_date.iloc[:, rcol], errors="coerce")
        lo, hi = remain_range
        df_date = df_date[(r > lo) & (r < hi)]

    df_date = df_date.replace(0, np.nan).dropna()
    x = pd.to_numeric(df_date.iloc[:, 0], errors="coerce")
    y = pd.to_numeric(df_date.iloc[:, 1], errors="coerce")
    if len(x) < 5:
        raise ValueError("insufficient points")

    popt, _ = curve_fit(inverse_cubic, x, y)
    a, b, c, d = popt
    formula = f"转股溢价率 = {a:.2f}/平价^3 + {b:.2f}/平价^2 + {c:.2f}/平价 + {d:.2f}"
    premium_100 = float(inverse_cubic(100, a, b, c, d))
    integral_value, _ = quad(lambda t: inverse_cubic(t, a, b, c, d), 70, 130)
    cal = integral_value / 60
    if with_greeks:
        return [
            "",
            a,
            b,
            c,
            d,
            premium_100,
            float(delta_cal(100, a, b, c, d)),
            float(gamma_cal(100, a, b, c, d)),
            formula,
            cal,
        ]
    return ["", a, b, c, d, premium_100, formula, cal]


def _run_fit_table(
    fit_data: SheetDict,
    *,
    floor_mode: FloorMode = None,
    remain_range: RemainMode = None,
    with_greeks: bool = False,
) -> pd.DataFrame:
    plain = _matrix(fit_data["平价"])
    prem = _matrix(fit_data["转股溢价率"])
    floor = _matrix(fit_data["平价底价溢价率"]) if floor_mode else None
    remain = _matrix(fit_data["剩余期限"]) if remain_range else None
    cols_main = ["日期", "a", "b", "c", "d", "转股溢价率", "拟合公式", "积分"]
    cols_greek = [
        "日期",
        "a",
        "b",
        "c",
        "d",
        "转股溢价率",
        "delta",
        "gamma",
        "拟合公式",
        "积分",
    ]
    columns = cols_greek if with_greeks else cols_main
    rows = []
    for d in date_cols(plain):
        try:
            row = _fit_one_date(
                plain[d],
                prem[d],
                floor[d] if floor is not None else None,
                floor_mode,
                remain[d] if remain is not None else None,
                remain_range,
                with_greeks=with_greeks,
            )
            row[0] = d
        except Exception:
            row = [d] + [0] * (len(columns) - 1)
        rows.append(row)
    return pd.DataFrame(rows, columns=columns)


def run_premium_fit(data: SheetDict, folder_name: str, mmdd: str) -> tuple[SheetDict, str]:
    fit_data = prepare_fit_data(data)
    result: SheetDict = {
        "百元平价拟合溢价率": _run_fit_table(fit_data, with_greeks=True),
        "平衡型": _run_fit_table(fit_data, floor_mode="balance"),
        "偏债型": _run_fit_table(fit_data, floor_mode="bond"),
        "偏股型": _run_fit_table(fit_data, floor_mode="stock"),
        "百元平价拟合溢价率（1-5.5年）": _run_fit_table(
            fit_data, remain_range=(1, 5.5)
        ),
        "百元平价拟合溢价率（0-2年）": _run_fit_table(
            fit_data, remain_range=(0, 2)
        ),
        "百元平价拟合溢价率（2-4年）": _run_fit_table(
            fit_data, remain_range=(2, 4)
        ),
        "百元平价拟合溢价率（4-6年）": _run_fit_table(
            fit_data, remain_range=(4, 6)
        ),
    }

    plot_dir = os.path.join(folder_name, f"{mmdd}日内估值数据更新")
    os.makedirs(plot_dir, exist_ok=True)
    _save_fit_plots(fit_data, result["百元平价拟合溢价率"], plot_dir)
    return result, plot_dir


def _save_fit_plots(
    fit_data: SheetDict, fit_main: pd.DataFrame, plot_dir: str
) -> None:
    plain = _matrix(fit_data["平价"])
    prem = _matrix(fit_data["转股溢价率"])
    dates = list(fit_main["日期"])
    if len(dates) < 1:
        return
    for d in (dates[-1], dates[0]) if len(dates) > 1 else (dates[-1],):
        row = fit_main.loc[fit_main["日期"] == d]
        if row.empty or float(row.iloc[0]["a"]) == 0:
            continue
        a, b, c, d0 = row.iloc[0][["a", "b", "c", "d"]]
        x = plain[d].dropna()
        if x.empty:
            continue
        y_act = prem[d].reindex(x.index)
        y_fit = inverse_cubic(x, a, b, c, d0)
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.scatter(x, y_act, s=10, alpha=0.5, label="实际")
        xs = np.linspace(max(50, x.min()), min(200, x.max()), 100)
        ax.plot(xs, inverse_cubic(xs, a, b, c, d0), "r-", label="拟合")
        ax.set_xlabel("平价")
        ax.set_ylabel("转股溢价率")
        ax.legend()
        ax.set_title(f"百元平价拟合 {d}")
        fig.savefig(
            os.path.join(plot_dir, f"{d}【华创固收】盘中百元平价拟合溢价.jpg"),
            dpi=120,
            bbox_inches="tight",
        )
        plt.close(fig)


def _apply_remain_duration_mask(data: SheetDict, min_years: float = 1.0) -> None:
    rem = data["剩余期限"]
    for d in date_cols(rem):
        bad = pd.to_numeric(rem[d], errors="coerce") < min_years
        for name, df in data.items():
            if name in ("总表", "发行规模") or d not in df.columns:
                continue
            df.loc[bad.values, d] = np.nan


def _apply_bond_type_ytm_mask(data: SheetDict) -> None:
    floor = data["平价底价溢价率"]
    for d in date_cols(floor):
        bad = pd.to_numeric(floor[d], errors="coerce") >= -20
        for name in ("余额", "YTM"):
            data[name].loc[bad.values, d] = np.nan


def run_js_update(
    data: SheetDict,
    stats_yao: SheetDict,
    folder_name: str,
    mmdd: str,
) -> SheetDict:
    """JS 基于 B 轨（剔妖）；平底/YTM 表剔除剩余期限<1年。"""
    js_data = prepare_masked_data(data, apply_track_b=True)
    colum = date_cols(js_data["余额"])

    need = ["平价", "纯债价值", "纯债溢价率", "隐含波动率"]
    lines = []
    for c in colum:
        lines.append([pd.to_numeric(js_data[n][c], errors="coerce").mean() for n in need])
    js: SheetDict = {
        "JS全样本": pd.DataFrame(
            lines,
            columns=[
                "平价算术平均",
                "纯债价值算术平均",
                "纯债溢价率算术平均",
                "隐含波动率算术平均",
            ],
            index=colum,
        )
    }

    tm1, tm2, tm3 = (
        _matrix(js_data["平价"]),
        _matrix(js_data["转股溢价率"]),
        _matrix(js_data["余额"]),
    )
    plines = []
    for i in colum:
        buckets = {k: [] for k in ["平价80以下", "平价80-95", "平价95-110", "平价110-125", "平价125以上"]}
        for a, b, c in zip(tm1[i], tm2[i], tm3[i]):
            if pd.isna(a) or pd.isna(b) or pd.isna(c):
                continue
            a = float(a)
            if a <= 80:
                buckets["平价80以下"].append((b, c))
            elif a <= 95:
                buckets["平价80-95"].append((b, c))
            elif a <= 110:
                buckets["平价95-110"].append((b, c))
            elif a <= 125:
                buckets["平价110-125"].append((b, c))
            else:
                buckets["平价125以上"].append((b, c))
        plines.append(tuple(avg_classify(buckets[k]) for k in buckets))
    js["平价分类转股溢价率"] = pd.DataFrame(
        plines,
        columns=["平价80以下", "平价80-95", "平价95-110", "平价110-125", "平价125以上"],
        index=colum,
    )

    js_floor = copy_sheet_dict(js_data)
    _apply_remain_duration_mask(js_floor, 1.0)
    tm1f = _matrix(js_floor["平价底价溢价率"])
    tm2f = _matrix(js_floor["转股溢价率"])
    tm3f = _matrix(js_floor["余额"])
    flines = []
    for i in colum:
        k1, k2, k3 = [], [], []
        for a, b, c in zip(tm1f[i], tm2f[i], tm3f[i]):
            if pd.isna(a) or pd.isna(b) or pd.isna(c):
                continue
            a = float(a)
            if a > 20:
                k1.append((b, c))
            elif a < -20:
                k2.append((b, c))
            else:
                k3.append((b, c))
        flines.append((avg_classify(k1), avg_classify(k2), avg_classify(k3)))
    js["JS平底分类余额加权转股溢价率"] = pd.DataFrame(
        flines, columns=["偏股", "偏债", "平衡"], index=colum
    )

    js_bond = copy_sheet_dict(js_floor)
    _apply_bond_type_ytm_mask(js_bond)
    bal = _matrix(js_bond["余额"])
    ytm = _matrix(js_bond["YTM"])
    ytm_bw, ytm_med = [], []
    for c in colum:
        ytm_bw.append([balance_weighted_mean(ytm[c], bal[c])])
        nums = [float(x) for x in ytm[c] if not pd.isna(x)]
        ytm_med.append([float(pd.Series(nums).quantile(0.5)) if nums else np.nan])
    js["JS偏债型余额YTM"] = pd.concat(
        [
            pd.DataFrame(ytm_bw, columns=["YTM余额加权"], index=colum),
            pd.DataFrame(ytm_med, columns=["YTM中位数"], index=colum),
        ],
        axis=1,
    )

    if "收盘价分位数统计" in stats_yao:
        js["收盘价分位数统计"] = stats_yao["收盘价分位数统计"].copy()

    return js


# ================== 行业均值与周报文本（内存） ==================

INDUSTRY_2021 = [
    "农林牧渔",
    "基础化工",
    "传媒",
    "电力设备",
    "电子",
    "房地产",
    "纺织服饰",
    "非银金融",
    "钢铁",
    "公用事业",
    "国防军工",
    "环保",
    "机械设备",
    "计算机",
    "家用电器",
    "建筑材料",
    "建筑装饰",
    "交通运输",
    "煤炭",
    "汽车",
    "轻工制造",
    "商贸零售",
    "社会服务",
    "石油石化",
    "食品饮料",
    "通信",
    "医药生物",
    "银行",
    "有色金属",
    "美容护理",
]
INDUSTRY_COL = "所属申万行业(2021）1级"
METRIC_SHEETS = ("收盘价", "转股溢价率", "平价", "纯债溢价率", "YTM")


def compute_industry_and_exclusion(data: SheetDict) -> SheetDict:
    """A 轨宽表上按日识别妖债并算申万一级行业算术均值。"""
    masked = copy_sheet_dict(data)
    apply_listing_delist_masks(masked)

    zb = masked["总表"]
    wide = {k: masked[k] for k in METRIC_SHEETS}
    cols = date_cols(wide["收盘价"])
    demo2: list[list] = [[] for _ in cols]
    close_df = wide["收盘价"]
    prem_df = wide["转股溢价率"]
    for j, d in enumerate(cols):
        col_i = j + 4
        names: list[str] = []
        for row in range(len(close_df)):
            cv = close_df.iloc[row, col_i]
            pv = prem_df.iloc[row, col_i]
            if pd.isna(cv):
                continue
            if float(cv) > TRACK_B_CLOSE_MIN and float(pv) > TRACK_B_PREMIUM_MIN:
                names.append(close_df.iloc[row, 0])
                for sheet in METRIC_SHEETS:
                    wide[sheet].iloc[row, col_i] = np.nan
        demo2[j] = names

    d1 = pd.DataFrame(demo2).T
    d1.columns = cols

    ind_means = {k: [] for k in METRIC_SHEETS}
    for ind in INDUSTRY_2021:
        idx = zb[INDUSTRY_COL] == ind
        for sheet in METRIC_SHEETS:
            sub = wide[sheet].loc[idx].iloc[:, 4:].apply(pd.to_numeric, errors="coerce")
            ind_means[sheet].append(sub.mean(numeric_only=True))

    out: SheetDict = {
        "剔除转债": d1,
    }
    sheet_cn = {
        "收盘价": "收盘价",
        "转股溢价率": "转股溢价率",
        "平价": "转换价值",
        "纯债溢价率": "纯债溢价率",
        "YTM": "YTM",
    }
    for sheet in METRIC_SHEETS:
        out[sheet_cn[sheet]] = pd.DataFrame(ind_means[sheet], index=INDUSTRY_2021, columns=cols)
    return out


def change_ratio(previous_value, current_value) -> str | None:
    if previous_value is None or current_value is None:
        return None
    if pd.isna(previous_value) or pd.isna(current_value) or previous_value == 0:
        return None
    rate = (current_value - previous_value) / previous_value
    if rate > 0:
        return f"上涨{abs(rate):.2%}"
    if rate < 0:
        return f"下跌{abs(rate):.2%}"
    return "维持同一水平"


def pct_change_ratio(change_ratio: float) -> str:
    if change_ratio > 0:
        return "上升" + str(round(change_ratio, 2)) + "pct"
    if change_ratio < 0:
        return "下降" + str(round(abs(change_ratio), 2)) + "pct"
    return "维持同一水平"


def _first_last(df: pd.DataFrame, col: str) -> tuple[float, float]:
    s = pd.to_numeric(df[col], errors="coerce")
    return float(s.iloc[0]), float(s.iloc[-1])


def append_weekly_report_sections(ctx: Any) -> None:
    """在已有 outputname 文件上续写统计/拟合段落（内存 DataFrame）。"""
    stats = ctx.stats_yao
    stats_clean = ctx.stats_clean
    fit = ctx.fit_result
    path = ctx.outputname

    full_bal = stats["全样本余额"]
    prev_close, cur_close = _first_last(full_bal, "收盘价余额加权")
    change_desc = change_ratio(prev_close, cur_close)

    with open(path, "a", encoding="utf-8") as f:
        f.write(f"\n转债收盘价加权平均值为{cur_close:.2f}，")
        if change_desc:
            f.write(f"较前周周五{change_desc}。")
        else:
            f.write("较前周周五数据不可用。")

    floor_close = stats["平底分类余额加权收盘价"]
    last_b, prev_b = _first_last(floor_close, "偏股")
    last_c, prev_c = _first_last(floor_close, "偏债")
    last_d, prev_d = _first_last(floor_close, "平衡")

    def _chg_desc(last, prev):
        if prev is None or last is None or prev == 0:
            return "未知"
        r = (last - prev) / prev
        if r > 0:
            return f"上升{abs(r):.2%}"
        if r < 0:
            return f"下降{abs(r):.2%}"
        return "维持同一水平"

    change_desc_b, change_desc_c, change_desc_d = (
        _chg_desc(last_b, prev_b),
        _chg_desc(last_c, prev_c),
        _chg_desc(last_d, prev_d),
    )

    interval = stats_clean["收盘价区间数量比例"]
    today = interval.iloc[-1]
    previous = interval.iloc[0]
    labels = list(interval.columns)
    today_high_ratio = float(today.iloc[6:8].sum() / today.sum())
    previous_high_ratio = float(previous.iloc[6:8].sum() / previous.sum())
    ratio_changed = (today_high_ratio - previous_high_ratio) * 100

    max_col, max_change_ratio, max_ratio = None, 0.0, 0.0
    for col in labels[3:6]:
        tr = float(today[col] / today.sum())
        pr = float(previous[col] / previous.sum())
        cr = (tr - pr) * 100
        if abs(cr) > abs(max_change_ratio):
            max_ratio, max_col, max_change_ratio = tr, col, cr

    low_count = float(today.iloc[0] + today.iloc[1] + today.iloc[2])

    with open(path, "a", encoding="utf-8") as f:
        f.write(
            f"其中偏股型转债的收盘价为{last_b:.2f}元，较上周五{change_desc_b}；"
            f"偏债型转债的收盘价为{last_c:.2f}元，较上周五{change_desc_c}；"
            f"平衡型转债的收盘价为{last_d:.2f}元，较上周五{change_desc_d}。"
        )
        f.write(
            f"\n从转债收盘价分布情况看，{max_col}区间占比"
            f"{'提升' if max_change_ratio > 0 else '下降' if max_change_ratio < 0 else '维持相同水平'}"
            f"较明显。"
        )
        f.write(
            f"截至{_time.strftime('%#m月%#d日', _time.localtime())}，"
            f"130元以上高价券个数占比{today_high_ratio:.2%}，"
            f"较上周五{pct_change_ratio(ratio_changed)}；"
        )
        f.write(
            f"占比变化最大的区间为{max_col}，占比{max_ratio:.2%}，"
            f"较上周五{pct_change_ratio(max_change_ratio)}；"
        )
        f.write(f"收盘价在100元以下的个券有{int(low_count)}只。\n")

    cq = stats["收盘价分位数统计"]
    med_last, med_prev = _first_last(cq, 0.5)
    med_chg = (med_last - med_prev) / med_prev if med_prev else 0
    with open(path, "a", encoding="utf-8") as f:
        f.write(
            f"价格中位数为{med_last:.2f}元，较前周周五"
            f"{'上升' if med_chg > 0 else '下降' if med_chg < 0 else '维持相同水平'}"
            f"{'' if med_chg == 0 else f'{abs(med_chg):.2%}'}。"
        )

    pct_cols = [0.05, 0.25, 0.75, 0.8, 0.9]
    parts = []
    for p in pct_cols:
        la, pr = _first_last(cq, p)
        ch = (la - pr) / pr if pr else 0
        parts.append(
            ("+" if ch > 0 else "-" if ch < 0 else "+") + str(round(ch * 100, 2)) + "%"
        )
    with open(path, "a", encoding="utf-8") as f:
        f.write(
            f"截至{_time.strftime('%#m月%#d日', _time.localtime())}，"
            f"5%、25%、75%、80%及90%分位数分别较前周五："
            + "、".join(parts)
            + "。\n"
        )

    fit_main = fit["百元平价拟合溢价率"]
    premium_last = float(fit_main.iloc[-1]["转股溢价率"])
    premium_prev = float(fit_main.iloc[0]["转股溢价率"])
    premium_delta = premium_last - premium_prev
    plain_last, plain_prev = _first_last(full_bal, "平价余额加权")
    plain_chg = (plain_last - plain_prev) / plain_prev if plain_prev else 0

    floor_prem = stats["平底分类余额加权转股溢价率"]
    stock_last, stock_prev = _first_last(floor_prem, floor_prem.columns[0])
    bond_last, bond_prev = _first_last(floor_prem, floor_prem.columns[1])
    bal_last, bal_prev = _first_last(floor_prem, floor_prem.columns[2])
    stock_d, bond_d, bal_d = stock_last - stock_prev, bond_last - bond_prev, bal_last - bal_prev

    with open(path, "a", encoding="utf-8") as f:
        f.write(
            f"\n转债市场百元平价拟合溢价率"
            f"{'抬升' if premium_delta > 0 else '压缩' if premium_delta < 0 else '维持此前水平'}。"
            f"截至{_time.strftime('%#m月%#d日', _time.localtime())}，"
            f"整体加权平价为{plain_last:.2f}元，较前周五"
            f"{'上升' if plain_chg > 0 else '下降' if plain_chg < 0 else '维持相同水平'}"
            f"{'' if plain_chg == 0 else f'{abs(plain_chg):.2%}'}。"
            f"百元平价拟合转股溢价率为{premium_last / 100:.2%}，较前周五"
            f"{'上升' if premium_delta > 0 else '下降' if premium_delta < 0 else '不变'}"
            f"{'' if premium_delta == 0 else f'{abs(premium_delta):.2f}pct'}；"
        )
        for label, last, delta in [
            ("偏股型", stock_last, stock_d),
            ("偏债型", bond_last, bond_d),
            ("平衡型", bal_last, bal_d),
        ]:
            f.write(
                f"{label}转债的转股溢价率为{last / 100:.2%}"
                f"{'，较前周周五上升' if delta > 0 else '，环比下降' if delta < 0 else '，环比维持相同水平'}"
                f"{'' if delta == 0 else f'{abs(delta):.2f}pct；'}"
            )
        f.write("\n")

    rating = stats["评级分类余额加权转股溢价率"]
    rating_cols = list(rating.columns)
    chg = [rating.iloc[-1][c] - rating.iloc[0][c] for c in rating_cols]
    with open(path, "a", encoding="utf-8") as f:
        f.write(
            f"\n截至{_time.strftime('%#m月%#d日', _time.localtime())}，"
            f"AAA评级{pct_change_ratio(chg[0])}，"
            f"AA+评级{pct_change_ratio(chg[1])}，"
            f"AA评级{pct_change_ratio(chg[2])}，"
            f"AA-评级{pct_change_ratio(chg[3])}，"
            f"A+评级{pct_change_ratio(chg[4])}。\n"
        )

    size = stats["规模分类余额加权转股溢价率"]
    size_cols = list(size.columns)
    schg = [size.iloc[-1][c] - size.iloc[0][c] for c in size_cols]
    size_names = [
        "50亿以上规模转债溢价率",
        "20-50亿（含50亿）规模",
        "10-20亿（含20亿）规模",
        "3-10亿（含10亿）规模",
        "3亿以下（含3亿）",
    ]
    with open(path, "a", encoding="utf-8") as f:
        f.write(
            f"从规模上看，{size_names[0]}{pct_change_ratio(schg[0])}，"
            f"{size_names[1]}{pct_change_ratio(schg[1])}，"
            f"{size_names[2]}{pct_change_ratio(schg[2])}，"
            f"{size_names[3]}{pct_change_ratio(schg[3])}，"
            f"{size_names[4]}{pct_change_ratio(schg[4])}。\n"
        )


def append_gf_daily_comment(
    ctx: Any,
    csi000832: pd.DataFrame,
    *,
    current_close_value: float,
    change_desc: str | None,
    premium_last: float,
    premium_delta: float,
) -> str:
    """GF 日评一行摘要（打印并可选写入）。"""
    last_date = str(ctx.last_date)
    start_date = str(ctx.start_date)
    amount_delta = (
        csi000832["成交额"][last_date] - csi000832["成交额"][start_date]
    )
    amount_delta_dir = "增量" if amount_delta > 0 else "缩量"
    how = (
        "大幅"
        if abs(amount_delta) > 100
        else ("小幅" if abs(amount_delta) < 30 else "")
    )
    price_delta = (
        csi000832["收盘价"][last_date] / csi000832["收盘价"][start_date] - 1
    )
    price_dir = "上涨" if price_delta > 0 else "下跌"
    line = (
        f"今日转债{how}{amount_delta_dir}{price_dir}，"
        f"成交额{csi000832['成交额'][last_date]:.2f}亿元，"
        f"环比{amount_delta_dir}{amount_delta:.2f}亿元，"
        f"中证转债{price_dir}{price_delta:.2%}，"
        f"转债收盘价加权平均值为{current_close_value:.2f}，"
        f"环比{change_desc or '—'}。"
        f"百元平价拟合转股溢价率为{premium_last / 100:.2%}，"
        f"环比{'上升' if premium_delta > 0 else '下降' if premium_delta < 0 else '不变'}"
        f"{'' if premium_delta == 0 else f'{abs(premium_delta):.2f}pct'}"
    )
    print(line)
    return line


# ================== 默认配置 ==================
IFIND_CREDENTIAL_FILE = Path(__file__).resolve().parents[2] / "private/ifind账号.txt"


def load_ifind_credentials() -> tuple[str, str]:
    """从项目目录的 ifind账号.txt 读取统一登录账号。"""
    if not IFIND_CREDENTIAL_FILE.is_file():
        raise FileNotFoundError(f"未找到iFinD账号文件：{IFIND_CREDENTIAL_FILE}")
    config = ConfigParser(interpolation=None)
    config.read(IFIND_CREDENTIAL_FILE, encoding="utf-8")
    username = config.get("ifind", "username", fallback="").strip()
    password = config.get("ifind", "password", fallback="").strip()
    if not username or not password:
        raise RuntimeError("ifind账号.txt中的[ifind] username或password为空")
    return username, password


INDEX_NAMES = ["000300.SH", "000905.SH", "000852.SH", "932000.CSI", "000832.CSI"]
INDEX_NAME_MAP = {
    "000300.SH": "上证综指",
    "000905.SH": "中证500",
    "000852.SH": "中证1000",
    "932000.CSI": "中证2000",
    "000832.CSI": "中证转债指数",
}
CB_BASIC_COLUMNS = [
    "转债简称",
    "正股代码",
    "正股简称",
    "发行方式",
    "交易状态",
    "转债余额",
    "上市日期",
]
EXCLUDED_CB_CODE = "128085.SZ"


# ================== 登录与自检 ==================
THS_LOGIN_OK_CODES = (0, -201)


def ths_login_errmsg(code: int) -> str:
    """获取 iFinD 登录/API 状态码对应的中文说明（与日报 notebook 一致）。"""
    try:
        info = THS_GetErrorInfo(code)
        if isinstance(info, dict):
            return str(info.get("errmsg", info))
        return str(info)
    except Exception:
        return f"未知错误（状态码 {code}）"


def is_ths_login_ok(code: int) -> bool:
    """0、-201 视为登录可用（-201 通常为已登录）。"""
    return code in THS_LOGIN_OK_CODES


def ths_login(ths_id: Optional[str] = None, ths_password: Optional[str] = None) -> int:
    """
    登录 iFinD，返回状态码。

    状态显示逻辑参考 THS【日报】API数据更新及周报文本：
    - 非 0/-201：失败，打印 errmsg
    - 0：成功，打印 THS_GetErrorInfo 返回的 errmsg（如 success!成功!）
    - -201：已登录，打印「登录成功！」
    """
    if not ths_id or not ths_password:
        file_id, file_password = load_ifind_credentials()
        ths_id = ths_id or file_id
        ths_password = ths_password or file_password
    code = THS_iFinDLogin(ths_id, ths_password)
    print(f"登录状态码: {code}")

    if code not in THS_LOGIN_OK_CODES:
        errmsg = ths_login_errmsg(code)
        print(f"登录失败: {errmsg}")
    elif code != -201:
        print(ths_login_errmsg(code))
    else:
        print("登录成功！")
    if is_ths_login_ok(code):
        print_data_statistics()
    return code


def thslogindemo(ths_id: Optional[str] = None, ths_password: Optional[str] = None) -> int:
    """兼容 notebook 命名。"""
    return ths_login(ths_id, ths_password)


def print_data_statistics() -> None:
    """打印 iFinD 各数据项配额使用占比。"""
    try:
        result = THS_DataStatistics()
        tables = result.get("tables", {}) if isinstance(result, dict) else {}
        if not tables:
            detail = result.get("errmsg", "未返回额度数据") if isinstance(result, dict) else str(result)
            print(f"[警告] iFinD使用额度查询失败：{detail}")
            return
        print("iFinD使用额度：")
        for key, value in tables.items():
            ratio = value.get("ratio", "N/A") if isinstance(value, dict) else value
            print(f"{key} 已用：{ratio}")
    except Exception as exc:
        print(f"[警告] iFinD使用额度查询失败：{exc}")


# ================== 日期与路径 ==================
def find_discontinuous_date_and_count(date_list: list[date]) -> tuple[date, int]:
    """从交易日序列末尾向前找第一个非连续间隔，返回间隔前一日及之后交易日个数。"""
    for i in range(len(date_list) - 1, 0, -1):
        if date_list[i] - date_list[i - 1] != timedelta(days=1):
            return date_list[i - 1], len(date_list) - i
    return date_list[0], len(date_list)


def resolve_date_range(
    days_today: int = 0,
    manual_days_backwards: int = -5,
) -> tuple[str, date, int, date]:
    """
    计算 last_date、start_date、days_backwards、上周最后交易日。

    参数与 notebook 一致：
    - days_today: 相对今日的交易日偏移（0=今天）
    - manual_days_backwards: 非 0 时覆盖自动推算的回溯天数
    """
    test_days_raw = THS_Date_Offset(
        "212001",
        "dateType:0,period:D,offset:-7,dateFormat:0,output:sequencedate",
        _time.strftime("%Y-%m-%d", _time.localtime()),
    ).data
    test_days = [datetime.strptime(d, "%Y-%m-%d").date() for d in test_days_raw.split(",")]
    discontinuous_date, count_from_discontinuous = find_discontinuous_date_and_count(test_days)

    if manual_days_backwards != 0:
        days_backwards = manual_days_backwards
    else:
        days_backwards = -count_from_discontinuous

    last_date = _time.strftime("%Y-%m-%d", _time.localtime())
    last_date = THS_Date_Offset(
        "212001",
        f"dateType:0,period:D,offset:{days_today},dateFormat:0,output:singledate",
        last_date,
    ).data

    days_raw = THS_Date_Offset(
        "212001",
        f"dateType:0,period:D,offset:{days_backwards},dateFormat:0,output:sequencedate",
        last_date,
    ).data
    days_list = [datetime.strptime(d, "%Y-%m-%d").date() for d in days_raw.split(",")]
    start_date = days_list[0]

    print(f"今天是 {last_date}")
    print(f"上周最后交易日为: {discontinuous_date}")
    print(f"全部区间为 {days_raw}")
    print(f"回溯天数: {days_backwards} 天")
    print(f"起始日为 {start_date}")
    return last_date, start_date, days_backwards, discontinuous_date


def setup_output_paths() -> tuple[str, str, str]:
    """创建 MMDD数据更新 目录，返回 folder_name、xlsx 路径、周报 txt 路径。"""
    mmdd = _time.strftime("%m%d", _time.localtime())
    folder_path = RUNS_DAILY_ROOT / f"{mmdd}数据更新"
    folder_path.mkdir(parents=True, exist_ok=True)
    folder_name = str(folder_path)
    filedir = str(folder_path / f"{mmdd}数据更新.xlsx")
    outputname = str(folder_path / f"{mmdd}转债周报.txt")
    print(f"更新数据保存路径为 {folder_name}")
    print("——————————————————————————————————————————————————————————————————————————")
    return folder_name, filedir, outputname


def _find_winrar_executable() -> str:
    candidates = (
        r"D:\WinRAR\Rar.exe",
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "WinRAR", "Rar.exe"),
        os.path.join(
            os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
            "WinRAR",
            "Rar.exe",
        ),
    )
    for path in candidates:
        if os.path.isfile(path):
            return path
    found = shutil.which("rar") or shutil.which("Rar")
    if found:
        return found
    raise RuntimeError("未找到 WinRAR（Rar.exe），请安装 WinRAR 或将其加入 PATH")


def archive_output_folder(folder_name: str) -> str:
    """将当日输出目录的内容打包为同级 RAR，不保留工作区父目录层级。"""
    abs_folder = os.path.abspath(folder_name)
    parent = os.path.dirname(abs_folder) or "."
    base = os.path.basename(abs_folder)
    archive_path = os.path.join(parent, f"{base}.rar")
    temp_archive_path = os.path.join(parent, f".{base}.tmp.rar")
    if os.path.isfile(temp_archive_path):
        os.remove(temp_archive_path)
    rar_exe = _find_winrar_executable()
    try:
        subprocess.run(
            [
                rar_exe,
                "a",
                "-r",
                "-dh",
                "-y",
                "-idq",
                "-x~$*",
                temp_archive_path,
                "*",
            ],
            check=True,
            cwd=abs_folder,
        )
        os.replace(temp_archive_path, archive_path)
    finally:
        if os.path.isfile(temp_archive_path):
            os.remove(temp_archive_path)
    print(f"已打包输出目录: {archive_path}")
    return archive_path


# ================== 文本格式化 ==================
def format_index_change(change: float) -> str:
    if change > 0:
        return f"上涨{abs(change / 100):.2%}"
    if change < 0:
        return f"下降{abs(change / 100):.2%}"
    return "持平"


def write_index_weekly_text(last_date: str, outputname: str) -> pd.DataFrame:
    """拉取指数周涨跌幅并写入周报 txt 首段。"""
    index_data = THS_BD(",".join(INDEX_NAMES), "ths_chg_ratio_w_index", str(last_date)).data
    index_data = index_data.set_index("thscode")

    descriptions = []
    for index_name in INDEX_NAMES:
        change = index_data.loc[index_name, "ths_chg_ratio_w_index"]
        name = INDEX_NAME_MAP.get(index_name, index_name)
        descriptions.append(f"{name}{format_index_change(change)}")

    text = "，".join(descriptions[:-1]) + "，" + descriptions[-1] + "。"
    with open(outputname, "w", encoding="utf-8") as f:
        f.write(text + "\n")
    return index_data


# ================== 转债列表与筛选 ==================
def fetch_cb_codes_on_date(edate_yyyymmdd: str) -> str:
    """可转债未到期列表，返回逗号分隔代码字符串。"""
    df = THS_DR(
        "p00570",
        f"jyzt=未到期;sfdb=全部;jysc=全部;edate={edate_yyyymmdd}",
        "jydm:Y,jydm_mc:Y,p00570_f001:Y,p00570_f019:Y",
        "format:dataframe",
    ).data
    return ",".join(df.set_index("jydm").index.astype(str))


def filter_cb_basic(df: pd.DataFrame) -> pd.DataFrame:
    """统一列名并剔除定向、NQ、终止上市等。"""
    out = df.set_index("thscode").rename_axis("转债代码")
    out.columns = CB_BASIC_COLUMNS
    out = out[~out["发行方式"].str.contains("定向", na=False)]
    out = out[~out.index.str.contains("NQ", na=False)]
    if "交易状态" in out.columns:
        out = out[~out["交易状态"].str.contains("终止上市", na=False)]
    return out


def fetch_cb_basic(codes: str, last_date: str) -> pd.DataFrame:
    raw = THS_BD(
        codes,
        "ths_convertible_debt_short_name_cbond;ths_stock_code_cbond;ths_stock_short_name_cbond;"
        "ths_issue_method_cbond;ths_trading_status_bond;ths_bond_balance_cbond;ths_listed_date_cbond",
        f";;;;;{last_date};",
    ).data
    return filter_cb_basic(raw)


def write_bonds_overview_text(cb_basic_all: pd.DataFrame, outputname: str) -> None:
    """写入全市场转债数量、余额及上市情况段落到周报 txt。"""
    index_count = len(cb_basic_all)
    total_balance = cb_basic_all["转债余额"].sum()

    unlisted = cb_basic_all[cb_basic_all["交易状态"] == "已发行未上市"]
    if len(unlisted) == 0:
        overview = (
            f"现已发行未到期可转债有{index_count}支，余额规模{total_balance:.2f}亿元，"
            "已发行转债中均已上市进行交易。"
        )
        with open(outputname, "a", encoding="utf-8") as f:
            f.write(overview)
        print(overview)
        return

    unlisted_names = "、".join(unlisted["转债简称"].astype(str))
    overview = (
        f"现已发行未到期可转债有{index_count}支，余额规模{total_balance:.2f}亿元，"
        f"已发行转债中，{unlisted_names}尚未上市进行交易，"
    )

    to_listed = unlisted[unlisted["上市日期"].notna()]
    to_listed_dict = {
        bond: datetime.strptime(str(date), "%Y%m%d").strftime("%m月%d日")
        for bond, date in to_listed[["转债简称", "上市日期"]].values
    }
    if len(to_listed_dict) == 1:
        k, v = next(iter(to_listed_dict.items()))
        listed_text = f"其中{k}将于{v}上市。"
    elif len(to_listed_dict) > 1:
        listed_text = (
            f"其中{'、'.join(to_listed_dict.keys())}将分别于"
            f"{'、'.join(to_listed_dict.values())}上市。"
        )
    else:
        listed_text = ""

    with open(outputname, "a", encoding="utf-8") as f:
        f.write(overview)
        f.write(listed_text)
    print(overview, listed_text)


def write_to_issue_bonds_text(outputname: str) -> pd.DataFrame | None:
    """即将发行转债段落。"""
    cb_list_to_issued = THS_DR(
        "p00600",
        "zqlx=640007",
        "p00600_f001:Y,p00600_f004:Y,p00600_f044:Y",
        "format:dataframe",
    ).data
    try:
        to_issued_dict = {
            bond: datetime.strptime(str(date), "%Y/%m/%d").strftime("%m月%d日")
            for bond, date in cb_list_to_issued[["p00600_f044", "p00600_f004"]].values
        }
        if len(to_issued_dict) == 1:
            k, v = next(iter(to_issued_dict.items()))
            text = f"其中{k}将于{v}网上发行。"
        else:
            text = (
                f"此外，{'、'.join(to_issued_dict.keys())}将分别于"
                f"{'、'.join(to_issued_dict.values())}网上发行。"
            )
        with open(outputname, "a", encoding="utf-8") as f:
            f.write(text)
        print(cb_list_to_issued)
        return cb_list_to_issued
    except Exception:
        msg = "目前尚无将发行转债。"
        with open(outputname, "a", encoding="utf-8") as f:
            f.write(msg)
        print(msg)
        return None


def build_trade_cb_universe(
    formatted_date_start: str,
    formatted_date_last: str,
    start_date: date,
    last_date: str,
) -> tuple[str, pd.DataFrame, pd.DataFrame]:
    """
    合并期初期末转债列表，按换手率剔除无交易券，返回可交易代码串与基础表。
    """
    start_list = THS_DR(
        "p00570",
        f"jyzt=未到期;sfdb=全部;jysc=全部;edate={formatted_date_start}",
        "jydm:Y,jydm_mc:Y,p00570_f001:Y,p00570_f019:Y",
        "format:dataframe",
    ).data.set_index("jydm").index.to_list()
    last_list = THS_DR(
        "p00570",
        f"jyzt=未到期;sfdb=全部;jysc=全部;edate={formatted_date_last}",
        "jydm:Y,jydm_mc:Y,p00570_f001:Y,p00570_f019:Y",
        "format:dataframe",
    ).data.set_index("jydm").index.to_list()

    cb_set = set(start_list + last_list)
    cb_set.discard(EXCLUDED_CB_CODE)
    cb_list = ", ".join(cb_set)

    trade_status = THS_DS(
        cb_list,
        "ths_turnover_ratio_cbond",
        "",
        "Fill:Blank,mode:thscode",
        str(start_date),
        str(last_date),
    ).data.set_index("time").T.iloc[:, [0, -1]]
    trade_status = trade_status[
        ~(
            (trade_status.iloc[:, 0].isna() & trade_status.iloc[:, -1].isna())
            | ((trade_status.iloc[:, 0] == 0) & (trade_status.iloc[:, -1] == 0))
        )
    ]

    cb_list_trade = ", ".join(trade_status.index.astype(str))
    cb_basic_trade = fetch_cb_basic(cb_list_trade, last_date)
    cb_list_trade = ", ".join(cb_basic_trade.index.astype(str))
    return cb_list_trade, cb_basic_trade, trade_status


# ================== 日度序列拉取 ==================
@dataclass
class CbTimeSeriesBundle:
    """区间内可交易转债的日度底层序列。"""

    ths_bond_balance_cbond: pd.DataFrame
    ths_bond_close_cbond: pd.DataFrame
    ths_specified_date_bond_rating_bond: pd.DataFrame
    ths_transfer_value_cbond: pd.DataFrame
    ths_conversion_premium_rate_cbond: pd.DataFrame
    ths_pure_bond_premium_rate_cbond: pd.DataFrame
    ths_pure_bond_value_cbond: pd.DataFrame
    ths_implied_volatility_cbond: pd.DataFrame
    issueamount: pd.DataFrame
    ths_pure_bond_ytm_cbond: pd.DataFrame
    ths_conversion_parity_price_premium_cbond: pd.DataFrame
    ths_turnover_ratio_cbond: pd.DataFrame
    ths_remain_duration_y_bond: pd.DataFrame
    ths_market_value_stock: pd.DataFrame
    ths_trading_status_bond: pd.DataFrame


def stock_list_from_cb_basic(cb_basic_trade: pd.DataFrame) -> str:
    """notebook 用正股代码拉 ths_market_value_stock，不能用转债代码。"""
    codes = cb_basic_trade["正股代码"].dropna().astype(str).unique()
    return ",".join(c for c in codes if c and c.lower() != "nan")


def fetch_cb_timeseries(
    cb_list_trade: str,
    stock_list: str,
    start_date: date,
    last_date: str,
) -> CbTimeSeriesBundle:

    s, e = str(start_date), str(last_date)

    def ds_fill(indicator: str, param: str = "") -> pd.DataFrame:
        return THS_DS(
            cb_list_trade, indicator, param, "Fill:Blank,mode:thscode", s, e
        ).data.set_index("time").T

    if stock_list.strip():
        stock_mv = (
            THS_DS(stock_list, "ths_market_value_stock", "", "mode:thscode", s, e)
            .data.set_index("time")
            .T
        )
    else:
        stock_mv = pd.DataFrame()

    return CbTimeSeriesBundle(
        ths_bond_balance_cbond=ds_fill("ths_bond_balance_cbond"),
        ths_bond_close_cbond=THS_DS(cb_list_trade, "ths_bond_close_cbond", "101", "Fill:Blank,mode:thscode", s, e)
        .data.set_index("time")
        .T,
        ths_specified_date_bond_rating_bond=THS_DS(
            cb_list_trade, "ths_specified_date_bond_rating_bond", "100", "mode:thscode", s, e
        )
        .data.set_index("time")
        .T,
        ths_transfer_value_cbond=ds_fill("ths_transfer_value_cbond"),
        ths_conversion_premium_rate_cbond=ds_fill("ths_conversion_premium_rate_cbond"),
        ths_pure_bond_premium_rate_cbond=ds_fill("ths_pure_bond_premium_rate_cbond"),
        ths_pure_bond_value_cbond=ds_fill("ths_pure_bond_value_cbond"),
        ths_implied_volatility_cbond=THS_DS(
            cb_list_trade, "ths_implied_volatility_cbond", "1", "Fill:Blank,mode:thscode", s, e
        )
        .data.set_index("time")
        .T,
        issueamount=THS_BD(cb_list_trade, "ths_issue_total_amt_cbond", "").data.set_index("thscode")
        / 100000000,
        ths_pure_bond_ytm_cbond=ds_fill("ths_pure_bond_ytm_cbond"),
        ths_conversion_parity_price_premium_cbond=THS_DS(
            cb_list_trade, "ths_conversion_parity_price_premium_cbond", "", "mode:thscode", s, e
        )
        .data.set_index("time")
        .T,
        ths_turnover_ratio_cbond=ds_fill("ths_turnover_ratio_cbond"),
        ths_remain_duration_y_bond=THS_DS(
            cb_list_trade, "ths_remain_duration_y_bond", "", "mode:thscode", s, e
        )
        .data.set_index("time")
        .T,
        ths_market_value_stock=stock_mv,
        ths_trading_status_bond=THS_DS(cb_list_trade, "ths_trading_status_bond", "", "mode:thscode", s, e)
        .data.set_index("time")
        .T,
    )



def _ths_extract_data(result: Any, *, api: str = "THS_BD", context: str = "") -> pd.DataFrame:
    label = api + (f"({context})" if context else "")
    if result is None:
        raise RuntimeError(f"{label} 返回 None")
    err = getattr(result, "errorcode", None)
    if err not in (None, 0) and err not in THS_LOGIN_OK_CODES:
        raise RuntimeError(f"{label} 失败: {ths_login_errmsg(int(err))}")
    data = getattr(result, "data", None)
    if data is None:
        extra = str(getattr(result, "errmsg", "") or "")
        raise RuntimeError(f"{label} data 为空（errorcode={err}）{extra}")
    return data


def _listing_date_param(listed: Any) -> str:
    dt = pd.to_datetime(listed, errors="coerce")
    if pd.isna(dt):
        raise ValueError(f"无效上市日期: {listed}")
    return dt.strftime("%Y-%m-%d")


def _fetch_ipo_close_series(codes: list[str], ipo: str) -> pd.Series:
    if not codes:
        return pd.Series(dtype=float)

    def _one_batch(batch: list[str]) -> pd.Series:
        raw = THS_BD(", ".join(batch), "ths_bond_close_cbond", f"{ipo},101")
        data = _ths_extract_data(raw, context=f"上市日收盘 {ipo}×{len(batch)}")
        df = data.set_index("thscode")
        col = "ths_bond_close_cbond" if "ths_bond_close_cbond" in df.columns else df.columns[0]
        return pd.to_numeric(df[col], errors="coerce")

    try:
        return _one_batch(codes)
    except RuntimeError as exc:
        if len(codes) == 1:
            print(f"警告: {exc}，代码 {codes[0]} 上市日价格置 NaN")
            return pd.Series([np.nan], index=codes)
        mid = len(codes) // 2
        left = _fetch_ipo_close_series(codes[:mid], ipo)
        right = _fetch_ipo_close_series(codes[mid:], ipo)
        return pd.concat([left, right])




def fetch_statistic_data(cb_list_trade: str, last_date: str) -> pd.DataFrame:
    """静态字段 + 按上市日批量取收盘价。"""
    raw_stat = THS_BD(
        cb_list_trade,
        "ths_listed_date_bond;ths_last_td_date_convertible_cbond;ths_conversion_sd_cbond;"
        "ths_object_the_sw_bond;ths_delist_date_bond",
        f";;;100,{last_date};",
    )
    stat = _ths_extract_data(raw_stat, context="转债静态字段").set_index("thscode")
    stat.columns = ["上市日期", "最后交易日", "转股期起始日", "申万行业", "摘牌日期"]

    listing_close = pd.Series(np.nan, index=stat.index, dtype="float64")
    for listed, group in tqdm(
        stat.dropna(subset=["上市日期"]).groupby("上市日期"),
        desc="上市日期收盘价",
    ):
        try:
            ipo = _listing_date_param(listed)
        except ValueError as exc:
            print(f"警告: {exc}，跳过该组")
            continue
        codes = group.index.astype(str).tolist()
        prices = _fetch_ipo_close_series(codes, ipo)
        listing_close.loc[group.index] = prices.reindex(group.index).values

    stat["上市日期价格"] = listing_close
    return stat


# ================== 内存宽表（不写回读） ==================

def _format_date_cell(val: Any) -> Any:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if len(s) == 8 and s.isdigit():
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    return s


def build_bond_meta(
    cb_basic_trade: pd.DataFrame,
    statistic_data: pd.DataFrame,
    rating_matrix: pd.DataFrame,
    last_date: str,
) -> pd.DataFrame:
    """宽表前四列：代码、名称、上市日期、评级（取期末评级）。"""
    codes = cb_basic_trade.index.astype(str)
    stat = statistic_data.reindex(codes)
    last_col = str(last_date)
    if last_col not in rating_matrix.columns and len(rating_matrix.columns):
        last_col = str(rating_matrix.columns[-1])
    rating = rating_matrix.reindex(codes)[last_col] if last_col in rating_matrix.columns else np.nan
    return pd.DataFrame(
        {
            "代码": codes,
            "名称": cb_basic_trade.reindex(codes)["转债简称"].values,
            "上市日期": stat["上市日期"].map(_format_date_cell).values,
            "评级": rating.values,
        }
    )


def build_stock_market_cap_wide(
    meta: pd.DataFrame,
    stock_matrix: pd.DataFrame,
    bond_to_stock: pd.Series,
) -> pd.DataFrame:
    """正股市值宽表：正股代码对齐到转债行，API 单位元 → 亿元（/1e8）。"""
    out = meta.copy()
    if stock_matrix.empty:
        return out
    sm = stock_matrix.apply(pd.to_numeric, errors="coerce") / 1e8
    bonds = meta["代码"].astype(str)
    stocks = bond_to_stock.reindex(bonds).astype(str)
    for col in sm.columns:
        col_vals = []
        for stk in stocks:
            if stk in ("", "nan", "None") or pd.isna(stk):
                col_vals.append(np.nan)
            elif stk in sm.index:
                col_vals.append(sm.at[stk, col])
            else:
                col_vals.append(np.nan)
        out[str(col)] = col_vals
    return out


def timeseries_to_wide(
    meta: pd.DataFrame, matrix: pd.DataFrame, *, numeric: bool = True
) -> pd.DataFrame:
    """日度矩阵（thscode × 日期）拼成 notebook 同款宽表。"""
    codes = meta["代码"].astype(str).tolist()
    m = matrix.reindex(codes)
    out = meta.copy()
    for col in m.columns:
        if numeric:
            out[str(col)] = pd.to_numeric(m[col], errors="coerce").values
        else:
            out[str(col)] = m[col].astype(object).where(m[col].notna(), None).values
    return out


def build_zongbiao_sheet(cb_basic_trade: pd.DataFrame, statistic_data: pd.DataFrame) -> pd.DataFrame:
    codes = cb_basic_trade.index.astype(str)
    stat = statistic_data.reindex(codes)
    conv = stat["转股期起始日"].map(_format_date_cell)
    return pd.DataFrame(
        {
            "代码": codes,
            "名称": cb_basic_trade.reindex(codes)["转债简称"].values,
            "上市日期": stat["上市日期"].map(_format_date_cell).values,
            "最后交易日": stat["最后交易日"].map(_format_date_cell).values,
            "转股期起始日": conv.values,
            "所属申万行业(2021）1级": stat["申万行业"].values,
            "摘牌日": stat["摘牌日期"].map(_format_date_cell).values,
            "上市首日价格": pd.to_numeric(stat["上市日期价格"], errors="coerce").values,
        }
    )


def build_issueamount_sheet(meta: pd.DataFrame, issueamount: pd.DataFrame, start_date: date) -> pd.DataFrame:
    codes = meta["代码"].astype(str).tolist()
    col = issueamount.columns[0]
    out = meta.copy()
    out[str(start_date)] = issueamount.reindex(codes)[col].values
    return out


def compute_subnew_wide_sheets(
    close_wide: pd.DataFrame,
    premium_wide: pd.DataFrame,
    zongbiao: pd.DataFrame,
    last_date: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    次新券指标仅内存计算，不单独落 raw xlsx。
    相对上市首日涨跌幅（%）、次新券转股溢价率（宽表同结构）。
    """
    meta = close_wide[list(WIDE_META_COLS)].copy()
    date_cols = [c for c in close_wide.columns if c not in WIDE_META_COLS]
    zb = zongbiao.set_index("代码")
    conv_start = pd.to_datetime(zb["转股期起始日"], errors="coerce")
    first_close = pd.to_numeric(zb["上市首日价格"], errors="coerce")
    last_dt = pd.to_datetime(last_date)

    close_m = close_wide.set_index("代码")[date_cols].apply(pd.to_numeric, errors="coerce")
    prem_m = premium_wide.set_index("代码")[date_cols].apply(pd.to_numeric, errors="coerce")
    # notebook：转股期起始日 > last_date 方为次新券；涨跌幅另需上市首日价
    is_subnew = conv_start > last_dt
    eligible_ipo = is_subnew & first_close.notna() & (first_close != 0)
    ipo_vals = close_m.div(first_close, axis=0).sub(1).mul(100)
    ipo_vals = ipo_vals.where(eligible_ipo.reindex(close_m.index), np.nan)
    # notebook 次新券转股溢价率：同上 is_subnew，仅填 eligible 券的溢价，非次新为空白
    prem_vals = prem_m.where(is_subnew.reindex(prem_m.index), np.nan)

    ipo_chg = meta.copy()
    prem = meta.copy()
    for dcol in date_cols:
        codes = meta["代码"].astype(str)
        ipo_chg[dcol] = ipo_vals[dcol].reindex(codes).values
        prem[dcol] = prem_vals[dcol].reindex(codes).values
    return ipo_chg, prem


def build_raw_data_sheets(ctx: "DailyUpdateContext") -> SheetDict:
    """由 API 结果组装中文宽表字典（唯一底稿来源，供后续计算读内存）。"""
    ts = ctx.timeseries
    meta = build_bond_meta(
        ctx.cb_basic_trade, ctx.statistic_data, ts.ths_specified_date_bond_rating_bond, ctx.last_date
    )
    implied_vol = ts.ths_implied_volatility_cbond.apply(pd.to_numeric, errors="coerce") * 100

    mapping = {
        "余额": ts.ths_bond_balance_cbond,
        "收盘价": ts.ths_bond_close_cbond,
        "债项评级": ts.ths_specified_date_bond_rating_bond,
        "平价": ts.ths_transfer_value_cbond,
        "转股溢价率": ts.ths_conversion_premium_rate_cbond,
        "纯债溢价率": ts.ths_pure_bond_premium_rate_cbond,
        "纯债价值": ts.ths_pure_bond_value_cbond,
        "隐含波动率": implied_vol,
        "YTM": ts.ths_pure_bond_ytm_cbond,
        "平价底价溢价率": ts.ths_conversion_parity_price_premium_cbond,
        "换手率": ts.ths_turnover_ratio_cbond,
        "剩余期限": ts.ths_remain_duration_y_bond,
    }
    raw: SheetDict = {}
    for name, mat in mapping.items():
        raw[name] = timeseries_to_wide(
            meta, mat, numeric=(name != "债项评级")
        )
    raw["正股市值"] = build_stock_market_cap_wide(
        meta,
        ts.ths_market_value_stock,
        ctx.cb_basic_trade["正股代码"],
    )
    raw["发行规模"] = build_issueamount_sheet(meta, ts.issueamount, ctx.start_date)
    raw["总表"] = build_zongbiao_sheet(ctx.cb_basic_trade, ctx.statistic_data)
    close_w = raw["收盘价"]
    prem_w = raw["转股溢价率"]
    sub_ipo, sub_pre = compute_subnew_wide_sheets(close_w, prem_w, raw["总表"], ctx.last_date)
    ctx.computed_wide["次新券相对上市首日涨跌幅"] = sub_ipo
    ctx.computed_wide["次新券转股溢价率"] = sub_pre
    return raw


def save_sheets_to_excel(
    sheets: SheetDict,
    path: str,
    *,
    index: bool = False,
    date_index: bool = False,
) -> None:
    """仅写出 Excel，不参与计算回路。

    date_index=True：按日期为行索引写出，首列标题为「日期」（与 notebook 汇总表一致）。
    """
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            if date_index:
                out = df.copy()
                if out.index.name is None:
                    out.index.name = "日期"
                out.to_excel(writer, sheet_name=name[:31], index=True)
            else:
                df.to_excel(writer, sheet_name=name[:31], index=index)


def export_cb_list(cb_basic_trade: pd.DataFrame, folder_name: str) -> str:
    mmdd = os.path.basename(folder_name).replace("数据更新", "")[:4]
    if not mmdd.isdigit():
        mmdd = _time.strftime("%m%d", _time.localtime())
    path = os.path.join(folder_name, f"{mmdd}可转债列表.xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        cb_basic_trade.to_excel(writer, sheet_name="可转债列表")
    return path


def data_for_statistics(raw: SheetDict, computed_wide: SheetDict) -> SheetDict:
    """汇总计算用的完整宽表视图（含仅内存的次新券列）。"""
    out = dict(raw)
    out.update(computed_wide)
    return out


# ================== 主流程 ==================
@dataclass
class DailyUpdateContext:
    """一次完整运行的内存上下文；下游步骤只读其字段，不读 Excel。"""

    folder_name: str
    filedir: str
    outputname: str
    last_date: str
    start_date: date
    days_backwards: int
    discontinuous_date: date
    index_data: pd.DataFrame
    cb_basic_all: pd.DataFrame
    cb_list_trade: str
    cb_basic_trade: pd.DataFrame
    trade_status: pd.DataFrame
    timeseries: CbTimeSeriesBundle
    statistic_data: pd.DataFrame
    cb_list_to_issued: pd.DataFrame | None = None
    raw_data: SheetDict = field(default_factory=dict)
    computed_wide: SheetDict = field(default_factory=dict)
    stats_clean: SheetDict = field(default_factory=dict)
    stats_yao: SheetDict = field(default_factory=dict)
    industry: SheetDict = field(default_factory=dict)
    fit_result: SheetDict = field(default_factory=dict)
    js_result: SheetDict = field(default_factory=dict)
    report_lines: list[str] = field(default_factory=list)
    cb_list_path: str = ""
    archive_path: str = ""
    elapsed_seconds: float = 0.0


def run_statistics_phase(ctx: DailyUpdateContext) -> None:
    """在内存中完成清理后 / 清理后剔妖汇总，并落盘统计 Excel（不回读）。"""
    full = data_for_statistics(ctx.raw_data, ctx.computed_wide)
    ctx.stats_clean = compute_statistics(full, apply_track_b=False)
    ctx.stats_yao = compute_statistics(full, apply_track_b=True)

    mmdd = os.path.basename(ctx.folder_name).replace("数据更新", "")[:4]
    if not mmdd.isdigit():
        mmdd = _time.strftime("%m%d", _time.localtime())
    path_clean = os.path.join(ctx.folder_name, f"{mmdd}数据更新（清理后）统计.xlsx")
    path_yao = os.path.join(ctx.folder_name, f"{mmdd}数据更新（清理后剔妖）统计.xlsx")
    save_sheets_to_excel(ctx.stats_clean, path_clean, date_index=True)
    save_sheets_to_excel(ctx.stats_yao, path_yao, date_index=True)
    print(f"已写出统计（内存计算）: {path_clean}")
    print(f"已写出统计（内存计算）: {path_yao}")


def _mmdd_from_folder(folder_name: str) -> str:
    mmdd = os.path.basename(folder_name).replace("数据更新", "")[:4]
    if not mmdd.isdigit():
        mmdd = _time.strftime("%m%d", _time.localtime())
    return mmdd


def run_fit_phase(ctx: DailyUpdateContext) -> None:
    full = data_for_statistics(ctx.raw_data, ctx.computed_wide)
    mmdd = _mmdd_from_folder(ctx.folder_name)
    ctx.fit_result, plot_dir = run_premium_fit(full, ctx.folder_name, mmdd)
    fit_path = os.path.join(ctx.folder_name, f"{mmdd}百元平价溢价率拟合结果.xlsx")
    save_sheets_to_excel(ctx.fit_result, fit_path, index=False)
    print(f"已写出拟合结果: {fit_path}")
    print(f"已保存拟合图: {plot_dir}")


def run_js_phase(ctx: DailyUpdateContext) -> None:
    full = data_for_statistics(ctx.raw_data, ctx.computed_wide)
    mmdd = _mmdd_from_folder(ctx.folder_name)
    ctx.js_result = run_js_update(full, ctx.stats_yao, ctx.folder_name, mmdd)
    js_path = os.path.join(ctx.folder_name, f"{mmdd}JS更新结果.xlsx")
    save_sheets_to_excel(ctx.js_result, js_path, date_index=True)
    print(f"已写出 JS 结果: {js_path}")


def run_industry_phase(ctx: DailyUpdateContext) -> None:
    full = data_for_statistics(ctx.raw_data, ctx.computed_wide)
    mmdd = _mmdd_from_folder(ctx.folder_name)
    ctx.industry = compute_industry_and_exclusion(full)
    ind_path = os.path.join(ctx.folder_name, f"{mmdd}剔除妖债及行业均值.xlsx")
    # 行业均值：行为申万行业、列为交易日；剔除转债：行为序号、列为交易日
    with pd.ExcelWriter(ind_path, engine="openpyxl") as writer:
        for name, df in ctx.industry.items():
            if name == "剔除转债":
                df.to_excel(writer, sheet_name=name[:31], index=False)
            else:
                out = df.copy()
                if out.index.name is None:
                    out.index.name = "行业"
                out.to_excel(writer, sheet_name=name[:31], index=True)
    print(f"已写出行业均值: {ind_path}")


def run_report_phase(ctx: DailyUpdateContext) -> None:
    append_weekly_report_sections(ctx)
    full_bal = ctx.stats_yao["全样本余额"]
    cur_close = float(full_bal["收盘价余额加权"].iloc[-1])
    prev_close = float(full_bal["收盘价余额加权"].iloc[0])
    change_desc = change_ratio(prev_close, cur_close)
    fit_main = ctx.fit_result["百元平价拟合溢价率"]
    premium_last = float(fit_main.iloc[-1]["转股溢价率"])
    premium_prev = float(fit_main.iloc[0]["转股溢价率"])
    csi = THS_DS(
        "000832.CSI",
        "ths_close_price_index;ths_trans_amt_index",
        ";",
        "block:history",
        str(ctx.start_date),
        str(ctx.last_date),
    ).data
    csi = csi.set_index("time")
    csi.columns = ["指数代码", "收盘价", "成交额"]
    csi["成交额"] = csi["成交额"] / 1e8
    append_gf_daily_comment(
        ctx,
        csi,
        current_close_value=cur_close,
        change_desc=change_desc,
        premium_last=premium_last,
        premium_delta=premium_last - premium_prev,
    )
    print(f"周报文本已续写: {ctx.outputname}")


def run_daily_data_update(
    days_today: int = 0,
    manual_days_backwards: int = -5,
    ths_id: Optional[str] = None,
    ths_password: Optional[str] = None,
    *,
    login: bool = True,
    show_quota: bool = True,
) -> DailyUpdateContext:
    """
    执行 notebook 中「登录 + 底层数据拉取」全流程。

    返回 DailyUpdateContext，内含各 DataFrame 与路径信息。
    """
    t0 = _time.time()
    if login:
        code = ths_login(ths_id, ths_password)
        if not is_ths_login_ok(code):
            raise RuntimeError(
                f"iFinD 登录失败（状态码 {code}）: {ths_login_errmsg(code)}"
            )

    folder_name, filedir, outputname = setup_output_paths()
    last_date, start_date, days_backwards, discontinuous_date = resolve_date_range(
        days_today, manual_days_backwards
    )

    index_data = write_index_weekly_text(last_date, outputname)

    formatted_date = datetime.strptime(last_date, "%Y-%m-%d").strftime("%Y%m%d")
    cb_codes_all = fetch_cb_codes_on_date(formatted_date)
    cb_basic_all = fetch_cb_basic(cb_codes_all, last_date)
    write_bonds_overview_text(cb_basic_all, outputname)
    cb_list_to_issued = write_to_issue_bonds_text(outputname)

    formatted_date_start = start_date.strftime("%Y%m%d")
    formatted_date_last = formatted_date
    cb_list_trade, cb_basic_trade, trade_status = build_trade_cb_universe(
        formatted_date_start, formatted_date_last, start_date, last_date
    )

    timeseries = fetch_cb_timeseries(
        cb_list_trade,
        stock_list_from_cb_basic(cb_basic_trade),
        start_date,
        last_date,
    )
    statistic_data = fetch_statistic_data(cb_list_trade, last_date)

    ctx = DailyUpdateContext(
        folder_name=folder_name,
        filedir=filedir,
        outputname=outputname,
        last_date=last_date,
        start_date=start_date,
        days_backwards=days_backwards,
        discontinuous_date=discontinuous_date,
        index_data=index_data,
        cb_basic_all=cb_basic_all,
        cb_list_trade=cb_list_trade,
        cb_basic_trade=cb_basic_trade,
        trade_status=trade_status,
        timeseries=timeseries,
        statistic_data=statistic_data,
        cb_list_to_issued=cb_list_to_issued,
    )
    ctx.raw_data = build_raw_data_sheets(ctx)
    save_sheets_to_excel(ctx.raw_data, filedir)
    ctx.cb_list_path = export_cb_list(cb_basic_trade, folder_name)
    print(f"已落盘底稿 {filedir}（计算使用内存 raw_data，不回读）")
    print(f"可转债列表 {ctx.cb_list_path}")

    run_statistics_phase(ctx)
    run_fit_phase(ctx)
    run_js_phase(ctx)
    run_industry_phase(ctx)
    run_report_phase(ctx)

    elapsed = _time.time() - t0
    print(f"拉取与汇总完成，耗时 {elapsed:.1f}s")
    ctx.elapsed_seconds = elapsed
    ctx.archive_path = archive_output_folder(folder_name)
    return ctx


def main() -> DailyUpdateContext:
    """命令行入口，参数与 notebook 默认一致。"""
    start_time = _time.time()
    ctx = run_daily_data_update(
        days_today=0,
        manual_days_backwards=-1,
    )
    print(f"总耗时 {_time.time() - start_time:.1f}s")
    return ctx


if __name__ == "__main__":
    main()
 
