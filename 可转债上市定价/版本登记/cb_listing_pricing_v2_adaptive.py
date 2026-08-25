from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import minimize


ROOT = Path(__file__).resolve().parents[2]
PROJECT_DIR = Path(__file__).resolve().parents[1]
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

BASE_HISTORY = ROOT / "tmp" / "cb_theme_revision" / "auto_theme_backtest_enhanced.csv"
ISSUE_FACTORS = ROOT / "tmp" / "cb_issue_factor_backtest" / "issue_factors.csv"
PRIOR_UPCOMING = ROOT / "tmp" / "cb_pricing_report_20260823" / "upcoming_forecasts.csv"
ANALYSIS_PATH = ROOT / "tmp" / "analyze_cb_listing_power_decay.py"


def import_analysis():
    spec = importlib.util.spec_from_file_location("cb_listing_analysis_v2", ANALYSIS_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


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


def load_history_with_latest_actual(analysis) -> pd.DataFrame:
    history = pd.read_csv(BASE_HISTORY, parse_dates=["上市日期", "预测信息日"])
    factors = pd.read_csv(ISSUE_FACTORS)
    history = history.merge(
        factors[["转债代码", "网上中签率", "原股东配售金额元"]],
        on="转债代码",
        how="left",
    )

    daily, _ = analysis.load_data()
    daily[analysis.DATE] = pd.to_datetime(daily[analysis.DATE]).dt.normalize()
    actual_lookup = daily.set_index([analysis.CODE, analysis.DATE])["收盘价"]
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


def build_current_targets(analysis, history: pd.DataFrame) -> tuple[pd.DataFrame, pd.Timestamp]:
    daily, master = analysis.load_data()
    daily = daily.merge(master[[analysis.CODE, "申万行业"]], on=analysis.CODE, how="left")
    dates = pd.DatetimeIndex(sorted(daily[analysis.DATE].dropna().unique()))
    latest = pd.Timestamp(dates.max())
    active = daily.loc[daily["转股溢价率"].notna() & daily["平价"].notna()].copy()
    active_by_date = {pd.Timestamp(d): g for d, g in active.groupby(analysis.DATE, sort=False)}

    def window_frame(window: int) -> pd.DataFrame:
        parts = [active_by_date.get(pd.Timestamp(d)) for d in dates[-window:]]
        return pd.concat([p for p in parts if p is not None], ignore_index=True)

    train1, train5, train10 = window_frame(1), window_frame(5), window_frame(10)
    meta = master.set_index(analysis.CODE)
    latest_rows = daily.loc[daily[analysis.DATE].eq(latest)].set_index(analysis.CODE)
    factors = pd.read_csv(ISSUE_FACTORS).set_index("转债代码")
    upcoming = master.loc[
        master["上市日期"].notna() & master["上市日期"].gt(latest)
    ].sort_values(["上市日期", analysis.CODE])
    rows = []
    for m in upcoming.itertuples(index=False):
        code = str(getattr(m, analysis.CODE))
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
        groups = analysis.classify_target(issue_size, market_cap, rating, industry)
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
        parts = [analysis.ridge_predict(analysis.prepare_fit(train1, parity), target_features, analysis.FULL_FEATURES, 0)]
        for group_type, label, frame in [
            ("板块", groups["板块"], train10),
            ("评级组", groups["评级组"], train5),
            ("余额组", groups["余额组"], train5),
            ("新券", "新券", train5),
            ("市值组", groups["市值组"], train5),
        ]:
            selected = frame.loc[analysis.group_mask(frame, group_type, label)].copy()
            parts.append(analysis.ridge_predict(analysis.prepare_fit(selected, parity), target_features, analysis.FULL_FEATURES, 0))
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


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-dir", default=str(PROJECT_DIR))
    args = parser.parse_args()
    project_dir = Path(args.project_dir)
    prediction_dir = project_dir / "预测"
    snapshot_dir = project_dir / "快照"
    version_dir = project_dir / "版本登记"
    for folder in [prediction_dir, snapshot_dir, version_dir]:
        folder.mkdir(parents=True, exist_ok=True)

    analysis = import_analysis()
    history = load_history_with_latest_actual(analysis)
    targets, latest = build_current_targets(analysis, history)
    raw = walk_forward_raw(history)

    forecasts = []
    selection_tables = []
    for _, target in targets.iterrows():
        selected, table = select_candidate(raw, pd.Timestamp(target["上市日期"]))
        forecast, calibrated_history = forecast_target(raw, target, selected)
        forecast["严格T-1"] = bool(pd.Timestamp(target["上市日期"]) == latest + pd.offsets.BDay(1))
        forecasts.append(forecast)
        table.insert(0, "目标转债代码", target["转债代码"])
        table.insert(1, "目标转债名称", target["转债名称"])
        table["是否入选"] = (
            table["logit_l2"].eq(selected["logit_l2"])
            & table["ridge_alpha"].eq(selected["ridge_alpha"])
            & table["probability_weight"].eq(selected["probability_weight"])
            & table["prob_cal_window"].eq(selected["prob_cal_window"])
            & table["price_cal_window"].eq(selected["price_cal_window"])
        )
        selection_tables.append(table)

    result = pd.DataFrame(forecasts)
    result.to_csv(prediction_dir / "v2_current_forecasts.csv", index=False, encoding="utf-8-sig")
    pd.concat(selection_tables, ignore_index=True).to_csv(snapshot_dir / "v2_current_candidate_selection.csv", index=False, encoding="utf-8-sig")
    history.to_csv(snapshot_dir / "v2_history_input_snapshot.csv", index=False, encoding="utf-8-sig")
    summary = {
        "model_version": "cb_listing_v2.0_adaptive",
        "data_as_of": latest.strftime("%Y-%m-%d"),
        "training_rule": "each target uses prior three years, post-rule samples only",
        "selection_rule": "minimum walk-forward MAE subject to RMSE and non-cap MAE no worse than fixed-weight baseline",
        "forecast_count": int(len(result)),
        "forecasts": result[
            [
                "转债代码", "转债名称", "上市日期", "预测信息日", "严格T-1",
                "最终预测价", "80%区间下限", "80%区间上限", "校准后触顶概率",
                "概率定价", "Ridge定价", "概率定价权重", "近三年内部样本外MAE",
                "近三年内部样本外RMSE", "近三年未触顶MAE",
            ]
        ].to_dict(orient="records"),
    }
    (version_dir / "v2_current_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
