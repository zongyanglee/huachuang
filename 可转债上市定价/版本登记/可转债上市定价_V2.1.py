from __future__ import annotations

"""可转债上市定价 v2.1 当前完整算法。

本文件固化当前正式预测链路，不导入项目内自定义模块。模型定义、参数网格、
数据口径及输出字段与 cb_listing_pricing_v21_scarcity.py 保持一致。
"""

import argparse
import json
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import curve_fit, minimize


ROOT = Path(__file__).resolve().parents[2]
PROJECT_DIR = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT / "data" / "转债个券历史序列"
DATA_START_YEAR = 2018

RULE_START = pd.Timestamp("2022-08-01")
CAP = 157.30
FLOOR = 56.70
LOOKBACK_YEARS = 3
MIN_TRAIN = 60
MIN_CLASS = 8

LOGIT_L2_GRID = [0.3, 1.0, 3.0, 10.0, 30.0]
RIDGE_ALPHA_GRID = [0.1, 1.0, 3.0, 10.0, 30.0, 100.0]
BLEND_GRID = sorted(set(np.round(np.arange(0, 1.0001, 0.05), 2).tolist() + [0.69]))
PROB_CAL_WINDOWS = [0, 20, 40, 60]
PRICE_CAL_WINDOWS = [0, 20, 40, 60]

CHIP_THRESHOLD_GRID = [0.50, 0.75, 1.00, 1.25]
CHIP_SLOPE_GRID = [6.0, 12.0, 18.0, 24.0]
CHIP_CAP_GRID = [6.0, 9.0, 12.0, 15.0]

BASE_HISTORY = ROOT / "tmp" / "cb_theme_revision" / "auto_theme_backtest_enhanced.csv"
ISSUE_FACTORS = ROOT / "tmp" / "cb_issue_factor_backtest" / "issue_factors.csv"
PRIOR_UPCOMING = ROOT / "tmp" / "cb_pricing_report_20260823" / "upcoming_forecasts.csv"

RESIDUAL_POWER_ANCHOR = 50.0
RESIDUAL_POWER_LOWER = 50.0
RESIDUAL_POWER_UPPER = 200.0
RESIDUAL_MIN_SAMPLES = 36
RESIDUAL_FEATURES = [
    "余额_log",
    "剩余期限",
    "正股20日波动率",
    "赎回累计天数",
    "下修累计天数",
    "隐含波动率",
]


def power_decay_with_floor(
    x: np.ndarray | float,
    amplitude: float,
    scale: float,
    power: float,
    floor: float,
    anchor_x: float = 70.0,
) -> np.ndarray | float:
    base = 1 + (np.asarray(x) - anchor_x) / scale
    return floor + amplitude * np.power(base, -power)


CODE = "转债代码"
DATE = "交易日期"
FULL_FEATURES = list(RESIDUAL_FEATURES)
EX_ANTE_FEATURES = ["余额_log", "剩余期限", "正股20日波动率"]
MIN_SAMPLES = RESIDUAL_MIN_SAMPLES
DEV_END = pd.Timestamp("2023-12-31")
VALIDATION_START = pd.Timestamp("2024-01-01")

SECTOR_GROUPS = {
    "科技": ("传媒", "电子", "国防军工", "计算机", "通信"),
    "金融": ("非银金融", "银行"),
    "制造": ("电力设备", "机械设备", "汽车", "轻工制造"),
    "消费": ("农林牧渔", "纺织服饰", "家用电器", "商贸零售", "社会服务", "食品饮料", "医药生物", "美容护理"),
    "周期": ("基础化工", "钢铁", "公用事业", "环保", "建筑材料", "建筑装饰", "交通运输", "煤炭", "石油石化", "有色金属"),
}
INDUSTRY_TO_SECTOR = {
    industry: sector for sector, industries in SECTOR_GROUPS.items() for industry in industries
}


@dataclass
class PreparedFit:
    work: pd.DataFrame
    base_pred: np.ndarray
    residual: np.ndarray
    base_target: float
    weights: np.ndarray


def load_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    columns = [
        CODE,
        DATE,
        "收盘价",
        "平价",
        "转股价",
        "转股溢价率",
        "换手率",
        "余额",
        "剩余期限",
        "正股20日波动率",
        "赎回累计天数",
        "下修累计天数",
        "隐含波动率",
        "债项评级",
        "正股市值",
        "正股收盘价",
    ]
    files = []
    for year_dir in sorted(DATA_ROOT.iterdir()):
        if not year_dir.is_dir() or not year_dir.name.isdigit() or int(year_dir.name) < DATA_START_YEAR:
            continue
        files.extend(sorted(year_dir.glob("*.parquet")))
    parts = []
    for idx, path in enumerate(files, 1):
        part = pd.read_parquet(path, columns=columns)
        part[DATE] = pd.to_datetime(part[DATE], errors="coerce").dt.normalize()
        parts.append(part)
        if idx % 12 == 0 or idx == len(files):
            print(f"[load] {idx}/{len(files)} {path.name}", flush=True)
    daily = pd.concat(parts, ignore_index=True)
    daily[CODE] = daily[CODE].astype(str)
    daily = daily.dropna(subset=[CODE, DATE])
    master = pd.read_parquet(DATA_ROOT / "_special" / "总表.parquet")
    master[CODE] = master[CODE].astype(str)
    master["上市日期"] = pd.to_datetime(master["上市日期"], errors="coerce").dt.normalize()
    return daily, master


def clean_training(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame
    work = frame.copy()
    for col in ["平价", "转股溢价率", "换手率", "余额", "剩余期限", "正股20日波动率", "赎回累计天数", "下修累计天数", "隐含波动率"]:
        work[col] = pd.to_numeric(work[col], errors="coerce")
    work = work.loc[
        work["平价"].gt(RESIDUAL_POWER_LOWER)
        & work["平价"].lt(RESIDUAL_POWER_UPPER)
        & work["换手率"].lt(50)
    ].copy()
    work = work.dropna(subset=["平价", "转股溢价率", "换手率"])
    if work.empty:
        return work
    low = float(work["转股溢价率"].quantile(0.03))
    high = float(work["转股溢价率"].quantile(0.97))
    work = work.loc[work["转股溢价率"].gt(low) & work["转股溢价率"].lt(high)].copy()
    work["余额_log"] = np.log1p(work["余额"].clip(lower=0))
    work = work.replace("", np.nan)
    for col in FULL_FEATURES:
        work.loc[work[col].eq(0), col] = np.nan
    return work


def prepare_fit(
    train: pd.DataFrame,
    target_parity: float,
    half_life_days: float | None = None,
) -> PreparedFit | None:
    work = clean_training(train)
    if len(work) < MIN_SAMPLES or not np.isfinite(target_parity):
        return None
    x = work["平价"].to_numpy(float)
    y = work["转股溢价率"].to_numpy(float)
    if half_life_days is None:
        weights = np.ones(len(work), dtype=float)
    else:
        ages = (work[DATE].max() - work[DATE]).dt.days.to_numpy(float)
        weights = np.power(0.5, ages / half_life_days)
        weights = np.clip(weights, 0.05, 1.0)
    floor0 = float(np.clip(np.nanpercentile(y, 5), 0, 1))
    amplitude0 = float(max(np.nanpercentile(y, 95) - floor0, 1))
    try:
        popt, _ = curve_fit(
            lambda xd, amplitude, scale, power, floor: power_decay_with_floor(
                xd, amplitude, scale, power, floor, anchor_x=RESIDUAL_POWER_ANCHOR
            ),
            x,
            y,
            p0=[amplitude0, 30.0, 2.0, floor0],
            bounds=([0, 1, 0.05, 0], [np.inf, 500, 20, 1]),
            sigma=1 / np.sqrt(weights),
            maxfev=30000,
        )
        base_pred = np.asarray(
            power_decay_with_floor(
                x, *map(float, popt), anchor_x=RESIDUAL_POWER_ANCHOR
            ),
            dtype=float,
        )
        base_target = float(
            power_decay_with_floor(
                target_parity, *map(float, popt), anchor_x=RESIDUAL_POWER_ANCHOR
            )
        )
        if not np.isfinite(base_target) or not np.isfinite(base_pred).all():
            return None
        return PreparedFit(work, base_pred, y - base_pred, base_target, weights)
    except Exception:
        return None


def ridge_predict(
    prepared: PreparedFit | None,
    target: dict,
    features: list[str],
    alpha: float,
) -> dict:
    empty = {"premium": np.nan, "price": np.nan, "n": 0, "rmse": np.nan, "base_premium": np.nan}
    if prepared is None:
        return empty
    work = prepared.work
    factors = work[features].apply(pd.to_numeric, errors="coerce").copy()
    target_values = []
    design_cols = [np.ones(len(work))]
    for feature in features:
        median = factors[feature].median()
        if pd.isna(median):
            median = 0.0
        filled = factors[feature].fillna(median)
        mean = float(filled.mean())
        std = float(filled.std(ddof=0))
        if not np.isfinite(std) or std <= 0:
            design_cols.append(np.zeros(len(work)))
            target_values.append(0.0)
        else:
            design_cols.append(((filled - mean) / std).to_numpy(float))
            raw_target = target.get(feature, np.nan)
            if raw_target is None or not np.isfinite(raw_target):
                raw_target = median
            target_values.append((float(raw_target) - mean) / std)
    X = np.column_stack(design_cols)
    xt = np.asarray([1.0, *target_values], dtype=float)
    sqrt_w = np.sqrt(prepared.weights)
    Xw = X * sqrt_w[:, None]
    yw = prepared.residual * sqrt_w
    penalty = np.eye(X.shape[1]) * float(alpha)
    penalty[0, 0] = 0.0
    try:
        beta = np.linalg.solve(Xw.T @ Xw + penalty, Xw.T @ yw)
    except np.linalg.LinAlgError:
        beta = np.linalg.lstsq(Xw.T @ Xw + penalty, Xw.T @ yw, rcond=None)[0]
    corrected = prepared.base_pred + X @ beta
    premium = float(prepared.base_target + xt @ beta)
    price = float(target["平价"] * (1 + premium / 100))
    rmse = float(np.sqrt(np.average((work["转股溢价率"].to_numpy(float) - corrected) ** 2, weights=prepared.weights)))
    return {
        "premium": premium,
        "price": price,
        "n": int(len(work)),
        "rmse": rmse,
        "base_premium": prepared.base_target,
    }


def classify_target(issue_size: float, market_cap: float, rating: str, industry: str) -> dict[str, str]:
    if issue_size < 3:
        balance_group = "0-3"
    elif issue_size < 10:
        balance_group = "3-10"
    elif issue_size < 20:
        balance_group = "10-20"
    elif issue_size < 50:
        balance_group = "20-50"
    else:
        balance_group = "50+"
    if market_cap < 50:
        cap_group = "0-50"
    elif market_cap < 300:
        cap_group = "50-300"
    else:
        cap_group = "300+"
    if rating in {"AAA", "AA+"}:
        rating_group = "AAA/AA+"
    elif rating in {"AA", "AA-"}:
        rating_group = "AA/AA-"
    elif rating in {"A+", "A"}:
        rating_group = "A+/A"
    else:
        rating_group = "其他"
    return {
        "板块": INDUSTRY_TO_SECTOR.get(industry, "其他"),
        "余额组": balance_group,
        "市值组": cap_group,
        "评级组": rating_group,
    }


def group_mask(frame: pd.DataFrame, group_type: str, label: str) -> pd.Series:
    if group_type == "板块":
        return frame["申万行业"].isin(SECTOR_GROUPS.get(label, ()))
    if group_type == "余额组":
        x = pd.to_numeric(frame["余额"], errors="coerce")
        bounds = {"0-3": (0, 3), "3-10": (3, 10), "10-20": (10, 20), "20-50": (20, 50)}
        if label == "50+":
            return x.gt(50)
        lo, hi = bounds[label]
        return x.gt(lo) & x.lt(hi)
    if group_type == "市值组":
        x = pd.to_numeric(frame["正股市值"], errors="coerce")
        if label == "0-50":
            return x.gt(0) & x.lt(50)
        if label == "50-300":
            return x.gt(50) & x.lt(300)
        return x.gt(300)
    if group_type == "评级组":
        values = {"AAA/AA+": {"AAA", "AA+"}, "AA/AA-": {"AA", "AA-"}, "A+/A": {"A+", "A"}}
        return frame["债项评级"].isin(values.get(label, set()))
    if group_type == "新券":
        x = pd.to_numeric(frame["剩余期限"], errors="coerce")
        return x.gt(5.5) & x.lt(6.01)
    return pd.Series(True, index=frame.index)

def sigmoid(x: np.ndarray | float) -> np.ndarray | float:
    return 1.0 / (1.0 + np.exp(-np.clip(x, -35, 35)))


def logit(x: np.ndarray | float) -> np.ndarray | float:
    x = np.clip(x, 0.01, 0.99)
    return np.log(x / (1.0 - x))


class Standardizer:
    def fit(self, frame: pd.DataFrame) -> "Standardizer":
        self.columns = list(frame.columns)
        work = frame.apply(pd.to_numeric, errors="coerce")
        self.medians = work.median().reindex(self.columns).fillna(0.0)
        filled = work.fillna(self.medians)
        self.lower = filled.quantile(0.01)
        self.upper = filled.quantile(0.99)
        clipped = filled.clip(self.lower, self.upper, axis=1)
        self.means = clipped.mean()
        self.stds = clipped.std(ddof=0).replace(0, 1).fillna(1)
        return self

    def transform(self, frame: pd.DataFrame) -> np.ndarray:
        work = frame.reindex(columns=self.columns).apply(pd.to_numeric, errors="coerce")
        work = work.fillna(self.medians).clip(self.lower, self.upper, axis=1)
        return ((work - self.means) / self.stds).to_numpy(float)


def fit_logistic(x: np.ndarray, y: np.ndarray, l2: float) -> np.ndarray:
    design = np.column_stack([np.ones(len(x)), x])

    def objective(beta: np.ndarray) -> tuple[float, np.ndarray]:
        z = design @ beta
        p = sigmoid(z)
        loss = np.sum(np.logaddexp(0, z) - y * z) + 0.5 * l2 * np.sum(beta[1:] ** 2)
        grad = design.T @ (p - y)
        grad[1:] += l2 * beta[1:]
        return float(loss), grad

    rate = np.clip(float(np.mean(y)), 1e-5, 1 - 1e-5)
    init = np.zeros(design.shape[1])
    init[0] = np.log(rate / (1 - rate))
    result = minimize(
        lambda b: objective(b)[0],
        init,
        jac=lambda b: objective(b)[1],
        method="L-BFGS-B",
    )
    return result.x


def predict_logistic(beta: np.ndarray, x: np.ndarray) -> np.ndarray:
    return sigmoid(np.column_stack([np.ones(len(x)), x]) @ beta)


def fit_ridge(x: np.ndarray, y: np.ndarray, alpha: float) -> np.ndarray:
    design = np.column_stack([np.ones(len(x)), x])
    penalty = np.eye(design.shape[1]) * float(alpha)
    penalty[0, 0] = 0.0
    return np.linalg.solve(design.T @ design + penalty, design.T @ y)


def predict_ridge(beta: np.ndarray, x: np.ndarray) -> np.ndarray:
    return np.column_stack([np.ones(len(x)), x]) @ beta


LOGIT_COLS = [
    "相似概率logit",
    "基础预测价",
    "近期20只新券触顶率",
    "ln发行规模",
    "中签热度",
    "原股东配售率",
]
RIDGE_COLS = ["ln发行规模", "中签热度", "原股东配售率"]


def derive_issue_features(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.copy()
    out["ln发行规模"] = np.log(pd.to_numeric(out["发行规模亿元"], errors="coerce").clip(lower=0.05))
    out["中签热度"] = -np.log(pd.to_numeric(out["网上中签率"], errors="coerce").clip(lower=1e-6))
    amount = pd.to_numeric(out["原股东配售金额元"], errors="coerce") / 1e8
    calculated_ratio = amount / pd.to_numeric(out["发行规模亿元"], errors="coerce")
    if "原股东配售率" not in out:
        out["原股东配售率"] = calculated_ratio
    else:
        out["原股东配售率"] = pd.to_numeric(
            out["原股东配售率"], errors="coerce"
        ).fillna(calculated_ratio)
    out["原股东配售率"] = pd.to_numeric(out["原股东配售率"], errors="coerce").clip(0, 1)
    peer = pd.to_numeric(out["动态相似新券触顶率"], errors="coerce")
    out["相似概率logit"] = logit(peer)
    if "实际上市收盘价" in out.columns:
        actual = pd.to_numeric(out["实际上市收盘价"], errors="coerce")
        out["实际触顶"] = actual.ge(CAP - 0.01).where(actual.notna())
    return out


def load_history_with_latest_actual() -> pd.DataFrame:
    history = pd.read_csv(BASE_HISTORY, parse_dates=["上市日期", "预测信息日"])
    factors = pd.read_csv(ISSUE_FACTORS)
    history = history.merge(
        factors[["转债代码", "网上中签率", "原股东配售金额元"]],
        on="转债代码",
        how="left",
    )

    daily, _ = load_data()
    daily[DATE] = pd.to_datetime(daily[DATE]).dt.normalize()
    actual_lookup = daily.set_index([CODE, DATE])["收盘价"]
    prior = pd.read_csv(PRIOR_UPCOMING, parse_dates=["上市日期", "预测信息日"])
    appended = []
    for row in prior.itertuples(index=False):
        key = (str(row.转债代码), pd.Timestamp(row.上市日期))
        if str(row.转债代码) in set(history["转债代码"].astype(str)) or key not in actual_lookup.index:
            continue
        actual = actual_lookup.loc[key]
        if isinstance(actual, pd.Series):
            actual = actual.iloc[-1]
        if not np.isfinite(actual):
            continue
        pre = history.loc[history["上市日期"].lt(row.上市日期)].sort_values("上市日期")
        recent20 = float(pre.tail(20)["实际触及157.30"].astype(float).mean())
        issue_row = factors.loc[factors["转债代码"].astype(str).eq(str(row.转债代码))]
        placement_amount = float(issue_row.iloc[-1]["原股东配售金额元"]) if len(issue_row) else np.nan
        appended.append({
            "转债代码": str(row.转债代码),
            "转债名称": row.转债名称,
            "上市日期": pd.Timestamp(row.上市日期),
            "预测信息日": pd.Timestamp(row.预测信息日),
            "实际上市收盘价": float(actual),
            "上市前一日平价": float(row.上市前一日平价),
            "发行规模亿元": float(row.发行规模亿元),
            "正股市值亿元": float(row.正股市值亿元),
            "正股20日波动率": float(row.正股20日波动率),
            "债项评级": row.债项评级,
            "申万行业": row.申万行业,
            "板块": row.板块,
            "基础预测价": float(row.基础预测价),
            "动态相似新券触顶率": float(row.动态相似新券触顶率),
            "动态相似新券预测价": float(row.动态相似新券预测价),
            "近期20只新券触顶率": recent20,
            "实际触及157.30": bool(float(actual) >= CAP - 0.01),
            "网上中签率": float(row.网上中签率),
            "原股东配售金额元": placement_amount,
            "原股东配售率": float(row.原股东配售率),
        })
    if appended:
        history = pd.concat([history, pd.DataFrame(appended)], ignore_index=True, sort=False)
    history = history.sort_values(["上市日期", "转债代码"]).drop_duplicates("转债代码", keep="last")
    history = history.loc[history["上市日期"].ge(RULE_START)].reset_index(drop=True)
    return derive_issue_features(history)


def dynamic_peer_probability(history: pd.DataFrame, target: pd.Series) -> float:
    prior = history.loc[
        history["上市日期"].lt(target["上市日期"])
        & history["上市日期"].ge(RULE_START)
    ].tail(100).copy()
    if len(prior) < 12:
        return float(prior["实际触顶"].mean()) if len(prior) else 0.0
    num = pd.DataFrame(index=prior.index)
    num["base"] = prior["基础预测价"]
    num["parity"] = prior["上市前一日平价"]
    num["issue"] = np.log1p(prior["发行规模亿元"].clip(lower=0))
    num["cap"] = np.log1p(prior["正股市值亿元"].clip(lower=0))
    num["vol"] = prior["正股20日波动率"]
    t = pd.Series({
        "base": target["基础预测价"],
        "parity": target["上市前一日平价"],
        "issue": np.log1p(max(target["发行规模亿元"], 0)),
        "cap": np.log1p(max(target["正股市值亿元"], 0)),
        "vol": target["正股20日波动率"],
    })
    med = num.median()
    scale = num.std(ddof=0).replace(0, 1).fillna(1)
    z = ((num.fillna(med) - t.fillna(med)) / scale).clip(-5, 5)
    dist = np.sqrt((z**2).mean(axis=1))
    rank_age = np.arange(len(prior) - 1, -1, -1)
    time_weight = np.power(0.5, rank_age / 60)
    same_ind = prior["申万行业"].eq(target["申万行业"]).to_numpy(float)
    same_sec = prior["板块"].eq(target["板块"]).to_numpy(float)
    weight = np.exp(-dist.to_numpy() / 0.75) * time_weight * (1 + 0.50 * same_ind + 0.15 * same_sec)
    best = np.argsort(weight)[-min(10, len(weight)):]
    w = weight[best]
    y = prior.iloc[best]["实际触顶"].to_numpy(float)
    cumulative = float(prior["实际触顶"].mean())
    return float((np.sum(w * y) + 3.0 * cumulative) / (np.sum(w) + 3.0))


def build_current_targets(history: pd.DataFrame) -> tuple[pd.DataFrame, pd.Timestamp]:
    daily, master = load_data()
    daily = daily.merge(master[[CODE, "申万行业"]], on=CODE, how="left")
    dates = pd.DatetimeIndex(sorted(daily[DATE].dropna().unique()))
    latest = pd.Timestamp(dates.max())
    active = daily.loc[daily["转股溢价率"].notna() & daily["平价"].notna()].copy()
    active_by_date = {pd.Timestamp(d): g for d, g in active.groupby(DATE, sort=False)}

    def window_frame(window: int) -> pd.DataFrame:
        parts = [active_by_date.get(pd.Timestamp(d)) for d in dates[-window:]]
        return pd.concat([p for p in parts if p is not None], ignore_index=True)

    train1, train5, train10 = window_frame(1), window_frame(5), window_frame(10)
    meta = master.set_index(CODE)
    latest_rows = daily.loc[daily[DATE].eq(latest)].set_index(CODE)
    factors = pd.read_csv(ISSUE_FACTORS).set_index("转债代码")
    upcoming = master.loc[
        master["上市日期"].notna() & master["上市日期"].gt(latest)
    ].sort_values(["上市日期", CODE])
    rows = []
    for m in upcoming.itertuples(index=False):
        code = str(getattr(m, CODE))
        if code not in latest_rows.index or code not in factors.index:
            continue
        r = latest_rows.loc[code]
        if isinstance(r, pd.DataFrame):
            r = r.iloc[-1]
        meta_row = meta.loc[code]
        conv = float(r["转股价"])
        stock = float(r["正股收盘价"])
        if not np.isfinite(conv) or conv <= 0 or not np.isfinite(stock) or stock <= 0:
            continue
        issue_size = float(meta_row["发行规模"])
        market_cap = float(r["正股市值"])
        vol = float(r["正股20日波动率"])
        industry = str(meta_row["申万行业"])
        rating_raw = r.get("债项评级", np.nan)
        rating = str(rating_raw) if pd.notna(rating_raw) else ""
        groups = classify_target(issue_size, market_cap, rating, industry)
        parity = stock / conv * 100
        target_features = {
            "平价": parity,
            "余额_log": float(np.log1p(issue_size)),
            "剩余期限": 6.0,
            "正股20日波动率": vol,
            "赎回累计天数": np.nan,
            "下修累计天数": np.nan,
            "隐含波动率": np.nan,
        }
        parts = [ridge_predict(prepare_fit(train1, parity), target_features, FULL_FEATURES, 0)]
        for group_type, label, frame in [
            ("板块", groups["板块"], train10),
            ("评级组", groups["评级组"], train5),
            ("余额组", groups["余额组"], train5),
            ("新券", "新券", train5),
            ("市值组", groups["市值组"], train5),
        ]:
            selected = frame.loc[group_mask(frame, group_type, label)].copy()
            parts.append(ridge_predict(prepare_fit(selected, parity), target_features, FULL_FEATURES, 0))
        valid = [x["price"] for x in parts if np.isfinite(x["price"])]
        if not valid:
            continue
        base = float(np.clip(np.median(valid), FLOOR, CAP))
        f = factors.loc[code]
        if isinstance(f, pd.DataFrame):
            f = f.iloc[-1]
        placement = float(f["原股东配售金额元"]) / 1e8 / issue_size
        target = pd.Series({
            "转债代码": code,
            "转债名称": meta_row["转债名称"],
            "上市日期": pd.Timestamp(meta_row["上市日期"]),
            "预测信息日": latest,
            "上市前一日平价": parity,
            "发行规模亿元": issue_size,
            "正股市值亿元": market_cap,
            "正股20日波动率": vol,
            "债项评级": rating,
            "申万行业": industry,
            "板块": groups["板块"],
            "基础预测价": base,
            "网上中签率": float(f["网上中签率"]),
            "原股东配售金额元": float(f["原股东配售金额元"]),
            "原股东配售率": float(np.clip(placement, 0, 1)),
        })
        peer_p = dynamic_peer_probability(history, target)
        prior = history.loc[history["上市日期"].lt(target["上市日期"])]
        target["动态相似新券触顶率"] = peer_p
        target["动态相似新券预测价"] = float(min(CAP, base + 0.75 * peer_p * max(CAP - base, 0)))
        target["近期20只新券触顶率"] = float(prior.tail(20)["实际触顶"].mean())
        rows.append(target.to_dict())
    return derive_issue_features(pd.DataFrame(rows)), latest


@dataclass
class RawWalkForward:
    frame: pd.DataFrame
    prob: dict[float, np.ndarray]
    ridge: dict[float, np.ndarray]


def walk_forward_raw(history: pd.DataFrame) -> RawWalkForward:
    data = history.sort_values(["上市日期", "转债代码"]).reset_index(drop=True).copy()
    prob = {l2: np.full(len(data), np.nan) for l2 in LOGIT_L2_GRID}
    ridge = {alpha: np.full(len(data), np.nan) for alpha in RIDGE_ALPHA_GRID}
    for i, row in data.iterrows():
        date = pd.Timestamp(row["上市日期"])
        start = max(date - pd.DateOffset(years=LOOKBACK_YEARS), RULE_START)
        train = data.loc[(data["上市日期"].ge(start)) & (data["上市日期"].lt(date))]
        if len(train) < MIN_TRAIN or train["实际触顶"].sum() < MIN_CLASS or (len(train) - train["实际触顶"].sum()) < MIN_CLASS:
            continue
        logit_scaler = Standardizer().fit(train[LOGIT_COLS])
        x_train = logit_scaler.transform(train[LOGIT_COLS])
        x_one = logit_scaler.transform(data.loc[[i], LOGIT_COLS])
        y = train["实际触顶"].to_numpy(float)
        for l2 in LOGIT_L2_GRID:
            beta = fit_logistic(x_train, y, l2)
            prob[l2][i] = float(predict_logistic(beta, x_one)[0])

        ridge_scaler = Standardizer().fit(train[RIDGE_COLS])
        xr_train = ridge_scaler.transform(train[RIDGE_COLS])
        xr_one = ridge_scaler.transform(data.loc[[i], RIDGE_COLS])
        residual = (train["实际上市收盘价"] - train["动态相似新券预测价"]).to_numpy(float)
        for alpha in RIDGE_ALPHA_GRID:
            beta = fit_ridge(xr_train, residual, alpha)
            correction = float(predict_ridge(beta, xr_one)[0])
            ridge[alpha][i] = float(np.clip(row["动态相似新券预测价"] + correction, FLOOR, CAP))
    return RawWalkForward(data, prob, ridge)


def calibrate_probability(
    raw_p: np.ndarray,
    y: np.ndarray,
    window: int,
    dates: np.ndarray | None = None,
) -> np.ndarray:
    out = raw_p.copy()
    if window <= 0:
        return out
    for i in range(len(raw_p)):
        if not np.isfinite(raw_p[i]):
            continue
        if dates is None:
            prior_mask = np.arange(len(raw_p)) < i
        else:
            prior_mask = dates < dates[i]
        prior = np.flatnonzero(np.isfinite(raw_p) & prior_mask)[-window:]
        if len(prior) < 8:
            continue
        obs_rate = (float(y[prior].sum()) + 1.0) / (len(prior) + 2.0)
        pred_rate = (float(raw_p[prior].sum()) + 1.0) / (len(prior) + 2.0)
        delta = float(logit(obs_rate) - logit(pred_rate))
        out[i] = float(sigmoid(logit(raw_p[i]) + delta))
    return out


def calibrate_price(
    raw_price: np.ndarray,
    actual: np.ndarray,
    window: int,
    dates: np.ndarray | None = None,
) -> np.ndarray:
    out = raw_price.copy()
    if window <= 0:
        return np.clip(out, FLOOR, CAP)
    residual = actual - raw_price
    for i in range(len(raw_price)):
        if not np.isfinite(raw_price[i]):
            continue
        if dates is None:
            prior_mask = np.arange(len(raw_price)) < i
        else:
            prior_mask = dates < dates[i]
        prior = np.flatnonzero(np.isfinite(residual) & prior_mask)[-window:]
        if len(prior) < 8:
            continue
        out[i] = raw_price[i] + float(np.median(residual[prior]))
    return np.clip(out, FLOOR, CAP)


def metric_values(actual: np.ndarray, pred: np.ndarray, hit: np.ndarray, mask: np.ndarray) -> dict:
    ok = mask & np.isfinite(actual) & np.isfinite(pred)
    a, p, h = actual[ok], pred[ok], hit[ok]
    err = p - a
    noncap = h == 0
    return {
        "n": int(len(a)),
        "mae": float(np.mean(np.abs(err))) if len(a) else np.nan,
        "rmse": float(np.sqrt(np.mean(err**2))) if len(a) else np.nan,
        "bias": float(np.mean(err)) if len(a) else np.nan,
        "within5": float(np.mean(np.abs(err) <= 5)) if len(a) else np.nan,
        "within10": float(np.mean(np.abs(err) <= 10)) if len(a) else np.nan,
        "noncap_n": int(noncap.sum()),
        "noncap_mae": float(np.mean(np.abs(err[noncap]))) if noncap.any() else np.nan,
    }


def candidate_predictions(raw: RawWalkForward, l2: float, alpha: float, blend: float, pwin: int, cwin: int) -> np.ndarray:
    data = raw.frame
    dates = data["上市日期"].to_numpy(dtype="datetime64[ns]")
    p = calibrate_probability(raw.prob[l2], data["实际触顶"].to_numpy(float), pwin, dates)
    base = data["基础预测价"].to_numpy(float)
    prob_price = np.minimum(CAP, base + p * np.maximum(CAP - base, 0))
    blended = blend * prob_price + (1 - blend) * raw.ridge[alpha]
    return calibrate_price(blended, data["实际上市收盘价"].to_numpy(float), cwin, dates)


def select_candidate(raw: RawWalkForward, target_date: pd.Timestamp) -> tuple[dict, pd.DataFrame]:
    data = raw.frame
    start = max(target_date - pd.DateOffset(years=LOOKBACK_YEARS), RULE_START)
    selection_mask = data["上市日期"].ge(start).to_numpy() & data["上市日期"].lt(target_date).to_numpy()
    actual = data["实际上市收盘价"].to_numpy(float)
    hit = data["实际触顶"].to_numpy(int)
    dates = data["上市日期"].to_numpy(dtype="datetime64[ns]")
    rows = []
    for l2 in LOGIT_L2_GRID:
        p_cache = {
            pwin: calibrate_probability(raw.prob[l2], hit.astype(float), pwin, dates)
            for pwin in PROB_CAL_WINDOWS
        }
        base = data["基础预测价"].to_numpy(float)
        for pwin, p in p_cache.items():
            prob_price = np.minimum(CAP, base + p * np.maximum(CAP - base, 0))
            for alpha in RIDGE_ALPHA_GRID:
                ridge_price = raw.ridge[alpha]
                for blend in BLEND_GRID:
                    uncalibrated = blend * prob_price + (1 - blend) * ridge_price
                    for cwin in PRICE_CAL_WINDOWS:
                        pred = calibrate_price(uncalibrated, actual, cwin, dates)
                        met = metric_values(actual, pred, hit, selection_mask)
                        if met["n"] < 60 or met["noncap_n"] < 8:
                            continue
                        rows.append({
                            "logit_l2": l2,
                            "ridge_alpha": alpha,
                            "probability_weight": float(blend),
                            "prob_cal_window": int(pwin),
                            "price_cal_window": int(cwin),
                            **met,
                        })
    table = pd.DataFrame(rows)
    if table.empty:
        raise RuntimeError("近三年内部时序预测不足，无法选择自适应模型")
    baseline = table.loc[
        table["logit_l2"].eq(3.0)
        & table["ridge_alpha"].eq(0.1)
        & table["probability_weight"].eq(0.69)
        & table["prob_cal_window"].eq(0)
        & table["price_cal_window"].eq(0)
    ].iloc[0]
    eligible = table.loc[
        table["rmse"].le(float(baseline["rmse"]) + 1e-12)
        & table["noncap_mae"].le(float(baseline["noncap_mae"]) + 1e-12)
    ].copy()
    if eligible.empty:
        selected = baseline.to_dict()
    else:
        eligible["complexity"] = (
            eligible["prob_cal_window"].gt(0).astype(int)
            + eligible["price_cal_window"].gt(0).astype(int)
        )
        selected = eligible.sort_values(
            ["mae", "noncap_mae", "rmse", "complexity", "probability_weight"],
            ascending=[True, True, True, True, False],
        ).iloc[0].to_dict()
    selected["baseline_mae"] = float(baseline["mae"])
    selected["baseline_rmse"] = float(baseline["rmse"])
    selected["baseline_noncap_mae"] = float(baseline["noncap_mae"])
    return selected, table


def probability_calibration_delta(raw_p: np.ndarray, y: np.ndarray, window: int) -> float:
    if window <= 0:
        return 0.0
    prior = np.flatnonzero(np.isfinite(raw_p))[-window:]
    if len(prior) < 8:
        return 0.0
    obs_rate = (float(y[prior].sum()) + 1.0) / (len(prior) + 2.0)
    pred_rate = (float(raw_p[prior].sum()) + 1.0) / (len(prior) + 2.0)
    return float(logit(obs_rate) - logit(pred_rate))


def forecast_target(raw: RawWalkForward, target: pd.Series, selected: dict) -> tuple[dict, np.ndarray]:
    date = pd.Timestamp(target["上市日期"])
    start = max(date - pd.DateOffset(years=LOOKBACK_YEARS), RULE_START)
    train = raw.frame.loc[(raw.frame["上市日期"].ge(start)) & (raw.frame["上市日期"].lt(date))].copy()

    l2 = float(selected["logit_l2"])
    alpha = float(selected["ridge_alpha"])
    blend = float(selected["probability_weight"])
    pwin = int(selected["prob_cal_window"])
    cwin = int(selected["price_cal_window"])

    logit_scaler = Standardizer().fit(train[LOGIT_COLS])
    beta_l = fit_logistic(logit_scaler.transform(train[LOGIT_COLS]), train["实际触顶"].to_numpy(float), l2)
    raw_p = float(predict_logistic(beta_l, logit_scaler.transform(pd.DataFrame([target])[LOGIT_COLS]))[0])
    delta = probability_calibration_delta(raw.prob[l2], raw.frame["实际触顶"].to_numpy(float), pwin)
    adj_p = float(sigmoid(logit(raw_p) + delta))
    base = float(target["基础预测价"])
    prob_price = float(min(CAP, base + adj_p * max(CAP - base, 0)))

    ridge_scaler = Standardizer().fit(train[RIDGE_COLS])
    beta_r = fit_ridge(
        ridge_scaler.transform(train[RIDGE_COLS]),
        (train["实际上市收盘价"] - train["动态相似新券预测价"]).to_numpy(float),
        alpha,
    )
    ridge_corr = float(predict_ridge(beta_r, ridge_scaler.transform(pd.DataFrame([target])[RIDGE_COLS]))[0])
    ridge_price = float(np.clip(float(target["动态相似新券预测价"]) + ridge_corr, FLOOR, CAP))
    uncalibrated = blend * prob_price + (1 - blend) * ridge_price

    historical_pred = candidate_predictions(raw, l2, alpha, blend, pwin, 0)
    residual = raw.frame["实际上市收盘价"].to_numpy(float) - historical_pred
    valid = np.flatnonzero(np.isfinite(residual) & raw.frame["上市日期"].lt(date).to_numpy())
    recent = valid[-cwin:] if cwin > 0 else np.array([], dtype=int)
    price_correction = float(np.median(residual[recent])) if len(recent) >= 8 else 0.0
    final = float(np.clip(uncalibrated + price_correction, FLOOR, CAP))

    calibrated_history = candidate_predictions(raw, l2, alpha, blend, pwin, cwin)
    history_error = np.abs(raw.frame["实际上市收盘价"].to_numpy(float) - calibrated_history)
    interval_idx = np.flatnonzero(np.isfinite(history_error) & raw.frame["上市日期"].lt(date).to_numpy())[-60:]
    radius = float(np.quantile(history_error[interval_idx], 0.8)) if len(interval_idx) >= 20 else np.nan

    return {
        **target.to_dict(),
        "训练样本数": int(len(train)),
        "训练触顶数": int(train["实际触顶"].sum()),
        "训练未触顶数": int(len(train) - train["实际触顶"].sum()),
        "选择_Logistic_L2": l2,
        "选择_Ridge_alpha": alpha,
        "概率定价权重": blend,
        "概率校准窗口": pwin,
        "价格校准窗口": cwin,
        "原始触顶概率": raw_p,
        "概率截距调整": delta,
        "校准后触顶概率": adj_p,
        "概率定价": prob_price,
        "Ridge修正额": ridge_corr,
        "Ridge定价": ridge_price,
        "价格校准额": price_correction,
        "最终预测价": final,
        "区间半径": radius,
        "80%区间下限": float(max(FLOOR, final - radius)) if np.isfinite(radius) else np.nan,
        "80%区间上限": float(min(CAP, final + radius)) if np.isfinite(radius) else np.nan,
        "近三年内部样本外MAE": float(selected["mae"]),
        "近三年内部样本外RMSE": float(selected["rmse"]),
        "近三年未触顶MAE": float(selected["noncap_mae"]),
        "旧固定权重基准MAE": float(selected["baseline_mae"]),
    }, calibrated_history

def candidate_matrix(raw):
    data = raw.frame
    actual = data["实际上市收盘价"].to_numpy(float)
    hit = data["实际触顶"].to_numpy(float)
    dates = data["上市日期"].to_numpy(dtype="datetime64[ns]")
    base = data["基础预测价"].to_numpy(float)
    base_rows = []
    uncalibrated_rows = []
    calibrated_prob = {
        (l2, pwin): calibrate_probability(raw.prob[l2], hit, pwin, dates)
        for l2 in LOGIT_L2_GRID
        for pwin in PROB_CAL_WINDOWS
    }
    for l2 in LOGIT_L2_GRID:
        for pwin in PROB_CAL_WINDOWS:
            p = calibrated_prob[(l2, pwin)]
            probability_price = np.minimum(CAP, base + p * np.maximum(CAP - base, 0))
            for alpha in RIDGE_ALPHA_GRID:
                ridge_price = raw.ridge[alpha]
                for blend in BLEND_GRID:
                    uncalibrated = blend * probability_price + (1 - blend) * ridge_price
                    base_rows.append((l2, alpha, float(blend), int(pwin)))
                    uncalibrated_rows.append(uncalibrated)

    uncalibrated_matrix = np.asarray(uncalibrated_rows, dtype=np.float64)

    def calibrate_matrix(window: int) -> np.ndarray:
        if window <= 0:
            return np.clip(uncalibrated_matrix, FLOOR, CAP)
        residual = actual[None, :] - uncalibrated_matrix
        out = uncalibrated_matrix.copy()
        for date in np.unique(dates):
            target_idx = np.flatnonzero(dates == date)
            prior_idx = np.flatnonzero(dates < date)
            if not len(prior_idx):
                continue
            valid_columns = prior_idx[np.isfinite(residual[0, prior_idx])][-window:]
            if len(valid_columns) < 8:
                continue
            correction = np.nanmedian(residual[:, valid_columns], axis=1)
            out[:, target_idx] += correction[:, None]
        return np.clip(out, FLOOR, CAP)

    calibrated = [calibrate_matrix(cwin) for cwin in PRICE_CAL_WINDOWS]
    predictions = np.stack(calibrated, axis=1).reshape(
        len(base_rows) * len(PRICE_CAL_WINDOWS), len(data)
    )
    rows = [
        (*base_row, int(cwin))
        for base_row in base_rows
        for cwin in PRICE_CAL_WINDOWS
    ]
    meta = pd.DataFrame(
        rows,
        columns=["logit_l2", "ridge_alpha", "probability_weight", "prob_cal_window", "price_cal_window"],
    )
    matrix = np.asarray(predictions, dtype=np.float32)
    baseline_mask = (
        meta["logit_l2"].eq(3.0)
        & meta["ridge_alpha"].eq(0.1)
        & meta["probability_weight"].eq(0.69)
        & meta["prob_cal_window"].eq(0)
        & meta["price_cal_window"].eq(0)
    )
    baseline_idx = int(np.flatnonzero(baseline_mask.to_numpy())[0])
    return meta, matrix, baseline_idx

def tradable_chips(frame: pd.DataFrame) -> np.ndarray:
    issue = pd.to_numeric(frame["发行规模亿元"], errors="coerce").to_numpy(float)
    placement = pd.to_numeric(frame["原股东配售率"], errors="coerce").to_numpy(float)
    return np.clip(issue * (1.0 - placement), 0.02, None)


def scarcity_adjustment(
    chips: np.ndarray | float,
    threshold: float,
    slope: float,
    correction_cap: float,
) -> np.ndarray:
    chips_array = np.asarray(chips, dtype=float)
    if slope <= 0 or threshold <= 0 or correction_cap <= 0:
        return np.zeros_like(chips_array)
    gap = np.maximum(0.0, np.log(threshold / np.clip(chips_array, 0.02, None)))
    return np.minimum(correction_cap, slope * gap)


def error_metrics(actual: np.ndarray, pred: np.ndarray, hit: np.ndarray) -> dict:
    ok = np.isfinite(actual) & np.isfinite(pred)
    actual = actual[ok]
    pred = pred[ok]
    hit = hit[ok]
    err = pred - actual
    noncap = ~hit
    return {
        "n": int(len(err)),
        "mae": float(np.mean(np.abs(err))),
        "rmse": float(np.sqrt(np.mean(err**2))),
        "bias": float(np.mean(err)),
        "within5": float(np.mean(np.abs(err) <= 5)),
        "within10": float(np.mean(np.abs(err) <= 10)),
        "under10": int(np.sum(err < -10)),
        "over10": int(np.sum(err > 10)),
        "noncap_n": int(noncap.sum()),
        "noncap_mae": float(np.mean(np.abs(err[noncap]))) if noncap.any() else np.nan,
    }


def select_scarcity_overlay(
    frame: pd.DataFrame,
    base_predictions: np.ndarray,
    target_date: pd.Timestamp,
    lookback_years: int = 3,
    minimum_samples: int = 60,
) -> tuple[dict, np.ndarray]:
    dates = pd.to_datetime(frame["上市日期"])
    start = target_date - pd.DateOffset(years=lookback_years)
    mask = dates.ge(start).to_numpy() & dates.lt(target_date).to_numpy()
    actual = pd.to_numeric(frame["实际上市收盘价"], errors="coerce").to_numpy(float)
    hit = actual >= 157.29
    chips = tradable_chips(frame)
    valid = mask & np.isfinite(actual) & np.isfinite(base_predictions) & np.isfinite(chips)
    baseline = error_metrics(actual[valid], base_predictions[valid], hit[valid])
    if baseline["n"] < minimum_samples:
        params = {
            "筹码阈值亿元": 0.0,
            "筹码斜率": 0.0,
            "单券修正上限": 0.0,
            "选择样本数": baseline["n"],
            "选择期原模型MAE": baseline["mae"],
            "选择期修正后MAE": baseline["mae"],
            "选择期原模型RMSE": baseline["rmse"],
            "选择期修正后RMSE": baseline["rmse"],
        }
        return params, base_predictions.copy()

    candidates = []
    zero = np.zeros(len(frame), dtype=float)
    candidates.append((baseline, 0.0, 0.0, 0.0, zero))
    for threshold in CHIP_THRESHOLD_GRID:
        for slope in CHIP_SLOPE_GRID:
            for correction_cap in CHIP_CAP_GRID:
                adjustment = scarcity_adjustment(chips, threshold, slope, correction_cap)
                corrected = np.minimum(157.30, base_predictions + adjustment)
                met = error_metrics(actual[valid], corrected[valid], hit[valid])
                if (
                    met["rmse"] <= baseline["rmse"] + 1e-12
                    and met["over10"] <= baseline["over10"]
                    and met["noncap_mae"] <= baseline["noncap_mae"] + 1e-12
                ):
                    candidates.append((met, threshold, slope, correction_cap, adjustment))

    met, threshold, slope, correction_cap, adjustment = min(
        candidates,
        key=lambda x: (x[0]["mae"], x[0]["rmse"], x[3], x[2], x[1]),
    )
    corrected_history = np.minimum(157.30, base_predictions + adjustment)
    params = {
        "筹码阈值亿元": float(threshold),
        "筹码斜率": float(slope),
        "单券修正上限": float(correction_cap),
        "选择样本数": int(baseline["n"]),
        "选择期原模型MAE": float(baseline["mae"]),
        "选择期修正后MAE": float(met["mae"]),
        "选择期原模型RMSE": float(baseline["rmse"]),
        "选择期修正后RMSE": float(met["rmse"]),
        "选择期原模型高估超10元": int(baseline["over10"]),
        "选择期修正后高估超10元": int(met["over10"]),
    }
    return params, corrected_history


def select_base_candidate(raw, meta, matrix, baseline_idx, target_date):
    data = raw.frame
    dates = pd.to_datetime(data["上市日期"])
    start = max(target_date - pd.DateOffset(years=LOOKBACK_YEARS), RULE_START)
    idx = np.flatnonzero(dates.ge(start).to_numpy() & dates.lt(target_date).to_numpy())
    actual = data["实际上市收盘价"].to_numpy(float)[idx]
    hit = data["实际触顶"].to_numpy(bool)[idx]
    pred = matrix[:, idx].astype(float)
    valid = np.isfinite(pred) & np.isfinite(actual)[None, :]
    n = valid.sum(axis=1)
    diff = pred - actual[None, :]
    mae = np.divide(np.where(valid, np.abs(diff), 0).sum(axis=1), n, out=np.full(len(meta), np.nan), where=n > 0)
    rmse = np.sqrt(np.divide(np.where(valid, diff**2, 0).sum(axis=1), n, out=np.full(len(meta), np.nan), where=n > 0))
    noncap_valid = valid & (~hit)[None, :]
    noncap_n = noncap_valid.sum(axis=1)
    noncap_mae = np.divide(
        np.where(noncap_valid, np.abs(diff), 0).sum(axis=1),
        noncap_n,
        out=np.full(len(meta), np.nan),
        where=noncap_n > 0,
    )
    eligible = (
        (n >= MIN_TRAIN)
        & (noncap_n >= MIN_CLASS)
        & (rmse <= rmse[baseline_idx] + 1e-12)
        & (noncap_mae <= noncap_mae[baseline_idx] + 1e-12)
    )
    eligible_idx = np.flatnonzero(eligible)
    if not len(eligible_idx):
        chosen_idx = baseline_idx
    else:
        complexity = (
            meta["prob_cal_window"].gt(0).astype(int)
            + meta["price_cal_window"].gt(0).astype(int)
        ).to_numpy()
        weight = meta["probability_weight"].to_numpy(float)
        order = np.lexsort((
            -weight[eligible_idx], complexity[eligible_idx], rmse[eligible_idx],
            noncap_mae[eligible_idx], mae[eligible_idx],
        ))
        chosen_idx = int(eligible_idx[order[0]])
    selected = meta.iloc[chosen_idx].to_dict()
    selected.update({
        "mae": float(mae[chosen_idx]),
        "rmse": float(rmse[chosen_idx]),
        "noncap_mae": float(noncap_mae[chosen_idx]),
        "baseline_mae": float(mae[baseline_idx]),
        "baseline_rmse": float(rmse[baseline_idx]),
        "baseline_noncap_mae": float(noncap_mae[baseline_idx]),
    })
    return selected

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--project-dir",
        default=str(PROJECT_DIR),
    )
    args = parser.parse_args()
    project_dir = Path(args.project_dir)
    prediction_dir = project_dir / "预测"
    version_dir = project_dir / "版本登记"
    prediction_dir.mkdir(parents=True, exist_ok=True)
    version_dir.mkdir(parents=True, exist_ok=True)

    history = load_history_with_latest_actual()
    targets, latest = build_current_targets(history)
    raw = walk_forward_raw(history)
    meta, matrix, baseline_idx = candidate_matrix(raw)

    rows = []
    for _, target in targets.iterrows():
        target_date = pd.Timestamp(target["上市日期"])
        selected = select_base_candidate(raw, meta, matrix, baseline_idx, target_date)
        base_forecast, base_history = forecast_target(raw, target, selected)
        scarcity, corrected_history = select_scarcity_overlay(raw.frame, base_history, target_date)
        target_chips = float(tradable_chips(pd.DataFrame([target]))[0])
        correction = float(scarcity_adjustment(
            target_chips,
            scarcity["筹码阈值亿元"],
            scarcity["筹码斜率"],
            scarcity["单券修正上限"],
        ))
        final = float(min(CAP, base_forecast["最终预测价"] + correction))
        actual = raw.frame["实际上市收盘价"].to_numpy(float)
        dates = pd.to_datetime(raw.frame["上市日期"])
        error = np.abs(actual - corrected_history)
        interval_idx = np.flatnonzero(
            np.isfinite(error) & dates.lt(target_date).to_numpy()
        )[-60:]
        radius = float(np.quantile(error[interval_idx], 0.8)) if len(interval_idx) >= 20 else np.nan
        rows.append({
            **base_forecast,
            "模型版本": "v2.1_极低筹码非线性修正",
            "潜在首日流通筹码亿元": target_chips,
            **scarcity,
            "v2.0预测价": float(base_forecast["最终预测价"]),
            "筹码非线性修正额": correction,
            "最终预测价": final,
            "区间半径": radius,
            "80%区间下限": float(max(FLOOR, final - radius)),
            "80%区间上限": float(min(CAP, final + radius)),
            "严格T-1": bool(target_date == latest + pd.offsets.BDay(1)),
        })

    result = pd.DataFrame(rows)
    result.to_csv(prediction_dir / "v21_current_forecasts.csv", index=False, encoding="utf-8-sig")
    summary = {
        "model_version": "cb_listing_v2.1_scarcity_overlay",
        "data_as_of": str(latest.date()),
        "scarcity_factor": "发行规模亿元 × (1 - 原股东配售率)",
        "scarcity_formula": "min(单券修正上限, 筹码斜率 × max(0, ln(筹码阈值/潜在首日流通筹码)))",
        "parameter_grids": {
            "threshold_billion_yuan": CHIP_THRESHOLD_GRID,
            "slope": CHIP_SLOPE_GRID,
            "correction_cap_yuan": CHIP_CAP_GRID,
        },
        "selection_constraints": "prior-three-year RMSE, overprediction-over-10 count and non-cap MAE must not be worse than v2.0",
        "forecast_count": int(len(result)),
        "forecasts": result[[
            "转债代码", "转债名称", "上市日期", "严格T-1",
            "潜在首日流通筹码亿元", "v2.0预测价", "筹码非线性修正额",
            "最终预测价", "80%区间下限", "80%区间上限",
            "筹码阈值亿元", "筹码斜率", "单券修正上限",
        ]].to_dict(orient="records"),
    }
    (version_dir / "v21_current_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()

