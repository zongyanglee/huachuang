from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
PROJECT_DIR = Path(__file__).resolve().parents[1]
V2_PATH = PROJECT_DIR / "版本登记" / "cb_listing_pricing_v2_adaptive.py"
V2_BACKTEST_PATH = PROJECT_DIR / "复核" / "backtest_cb_listing_pricing_v2_adaptive.py"

CHIP_THRESHOLD_GRID = [0.50, 0.75, 1.00, 1.25]
CHIP_SLOPE_GRID = [6.0, 12.0, 18.0, 24.0]
CHIP_CAP_GRID = [6.0, 9.0, 12.0, 15.0]


def import_file(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


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


def select_base_candidate(v2, raw, meta, matrix, baseline_idx, target_date):
    data = raw.frame
    dates = pd.to_datetime(data["上市日期"])
    start = max(target_date - pd.DateOffset(years=v2.LOOKBACK_YEARS), v2.RULE_START)
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
        (n >= v2.MIN_TRAIN)
        & (noncap_n >= v2.MIN_CLASS)
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

    v2 = import_file("cb_listing_pricing_v2_base", V2_PATH)
    bt = import_file("cb_listing_pricing_v2_backtest_helpers", V2_BACKTEST_PATH)
    analysis = v2.import_analysis()
    history = v2.load_history_with_latest_actual(analysis)
    targets, latest = v2.build_current_targets(analysis, history)
    raw = v2.walk_forward_raw(history)
    meta, matrix, baseline_idx = bt.candidate_matrix(v2, raw)

    rows = []
    for _, target in targets.iterrows():
        target_date = pd.Timestamp(target["上市日期"])
        selected = select_base_candidate(v2, raw, meta, matrix, baseline_idx, target_date)
        base_forecast, base_history = v2.forecast_target(raw, target, selected)
        scarcity, corrected_history = select_scarcity_overlay(raw.frame, base_history, target_date)
        target_chips = float(tradable_chips(pd.DataFrame([target]))[0])
        correction = float(scarcity_adjustment(
            target_chips,
            scarcity["筹码阈值亿元"],
            scarcity["筹码斜率"],
            scarcity["单券修正上限"],
        ))
        final = float(min(v2.CAP, base_forecast["最终预测价"] + correction))
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
            "80%区间下限": float(max(v2.FLOOR, final - radius)),
            "80%区间上限": float(min(v2.CAP, final + radius)),
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
