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
MODEL_PATH = PROJECT_DIR / "版本登记" / "cb_listing_pricing_v2_adaptive.py"


def import_model():
    spec = importlib.util.spec_from_file_location("cb_listing_pricing_v2", MODEL_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def stats(actual: np.ndarray, pred: np.ndarray) -> dict:
    ok = np.isfinite(actual) & np.isfinite(pred)
    err = pred[ok] - actual[ok]
    ae = np.abs(err)
    return {
        "n": int(ok.sum()),
        "mae": float(np.mean(ae)),
        "rmse": float(np.sqrt(np.mean(err**2))),
        "median_ae": float(np.median(ae)),
        "bias": float(np.mean(err)),
        "within_3": float(np.mean(ae <= 3)),
        "within_5": float(np.mean(ae <= 5)),
        "within_10": float(np.mean(ae <= 10)),
        "overprediction_share": float(np.mean(err > 0)),
        "underprediction_share": float(np.mean(err < 0)),
    }


def candidate_matrix(model, raw):
    data = raw.frame
    actual = data["实际上市收盘价"].to_numpy(float)
    hit = data["实际触顶"].to_numpy(float)
    dates = data["上市日期"].to_numpy(dtype="datetime64[ns]")
    base = data["基础预测价"].to_numpy(float)
    base_rows = []
    uncalibrated_rows = []
    calibrated_prob = {
        (l2, pwin): model.calibrate_probability(raw.prob[l2], hit, pwin, dates)
        for l2 in model.LOGIT_L2_GRID
        for pwin in model.PROB_CAL_WINDOWS
    }
    for l2 in model.LOGIT_L2_GRID:
        for pwin in model.PROB_CAL_WINDOWS:
            p = calibrated_prob[(l2, pwin)]
            probability_price = np.minimum(model.CAP, base + p * np.maximum(model.CAP - base, 0))
            for alpha in model.RIDGE_ALPHA_GRID:
                ridge_price = raw.ridge[alpha]
                for blend in model.BLEND_GRID:
                    uncalibrated = blend * probability_price + (1 - blend) * ridge_price
                    base_rows.append((l2, alpha, float(blend), int(pwin)))
                    uncalibrated_rows.append(uncalibrated)

    uncalibrated_matrix = np.asarray(uncalibrated_rows, dtype=np.float64)

    def calibrate_matrix(window: int) -> np.ndarray:
        if window <= 0:
            return np.clip(uncalibrated_matrix, model.FLOOR, model.CAP)
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
        return np.clip(out, model.FLOOR, model.CAP)

    calibrated = [calibrate_matrix(cwin) for cwin in model.PRICE_CAL_WINDOWS]
    predictions = np.stack(calibrated, axis=1).reshape(
        len(base_rows) * len(model.PRICE_CAL_WINDOWS), len(data)
    )
    rows = [
        (*base_row, int(cwin))
        for base_row in base_rows
        for cwin in model.PRICE_CAL_WINDOWS
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


def select_outer(model, raw, meta, matrix, baseline_idx):
    data = raw.frame.reset_index(drop=True)
    actual = data["实际上市收盘价"].to_numpy(float)
    hit = data["实际触顶"].to_numpy(bool)
    dates = pd.to_datetime(data["上市日期"])
    complexity = (
        meta["prob_cal_window"].gt(0).astype(int)
        + meta["price_cal_window"].gt(0).astype(int)
    ).to_numpy()
    probability_weight = meta["probability_weight"].to_numpy(float)
    records = []

    valid_all = np.isfinite(matrix) & np.isfinite(actual)[None, :]
    diff_all = np.where(valid_all, matrix.astype(np.float64) - actual[None, :], 0.0)
    abs_all = np.abs(diff_all)
    sq_all = diff_all**2
    noncap_all = valid_all & (~hit)[None, :]

    def cumulative(values, dtype=None):
        running = np.cumsum(values, axis=1, dtype=dtype)
        return np.concatenate([np.zeros((len(meta), 1), dtype=running.dtype), running], axis=1)

    count_cum = cumulative(valid_all, dtype=np.int32)
    noncap_count_cum = cumulative(noncap_all, dtype=np.int32)
    abs_cum = cumulative(np.where(valid_all, abs_all, 0.0), dtype=np.float64)
    sq_cum = cumulative(np.where(valid_all, sq_all, 0.0), dtype=np.float64)
    noncap_abs_cum = cumulative(np.where(noncap_all, abs_all, 0.0), dtype=np.float64)
    date_values = dates.to_numpy(dtype="datetime64[ns]")

    for i in range(len(data)):
        target_date = dates.iloc[i]
        start = max(target_date - pd.DateOffset(years=model.LOOKBACK_YEARS), model.RULE_START)
        left = int(np.searchsorted(date_values, np.datetime64(start), side="left"))
        right = int(np.searchsorted(date_values, np.datetime64(target_date), side="left"))
        if right <= left:
            continue
        n = count_cum[:, right] - count_cum[:, left]
        noncap_n = noncap_count_cum[:, right] - noncap_count_cum[:, left]
        if n[baseline_idx] < model.MIN_TRAIN or noncap_n[baseline_idx] < model.MIN_CLASS:
            continue

        abs_sum = abs_cum[:, right] - abs_cum[:, left]
        sq_sum = sq_cum[:, right] - sq_cum[:, left]
        noncap_abs_sum = noncap_abs_cum[:, right] - noncap_abs_cum[:, left]
        mae = np.divide(
            abs_sum, n,
            out=np.full(len(meta), np.nan), where=n > 0,
        )
        rmse = np.sqrt(np.divide(
            sq_sum, n,
            out=np.full(len(meta), np.nan), where=n > 0,
        ))
        noncap_mae = np.divide(
            noncap_abs_sum, noncap_n,
            out=np.full(len(meta), np.nan), where=noncap_n > 0,
        )
        eligible = (
            (n >= model.MIN_TRAIN)
            & (noncap_n >= model.MIN_CLASS)
            & (rmse <= rmse[baseline_idx] + 1e-12)
            & (noncap_mae <= noncap_mae[baseline_idx] + 1e-12)
        )
        eligible_idx = np.flatnonzero(eligible)
        if not len(eligible_idx):
            selected_idx = baseline_idx
        else:
            order = np.lexsort((
                -probability_weight[eligible_idx],
                complexity[eligible_idx],
                rmse[eligible_idx],
                noncap_mae[eligible_idx],
                mae[eligible_idx],
            ))
            selected_idx = int(eligible_idx[order[0]])

        chosen = meta.iloc[selected_idx]
        target_pred = float(matrix[selected_idx, i])
        baseline_pred = float(matrix[baseline_idx, i])
        if not np.isfinite(target_pred) or not np.isfinite(baseline_pred):
            continue
        records.append({
            **data.iloc[i].to_dict(),
            "自适应预测价": target_pred,
            "固定权重基准预测价": baseline_pred,
            "自适应误差": target_pred - actual[i],
            "自适应绝对误差": abs(target_pred - actual[i]),
            "固定权重基准误差": baseline_pred - actual[i],
            "选择依据样本数": int(n[selected_idx]),
            "选择依据未触顶样本数": int(noncap_n[selected_idx]),
            "选择期MAE": float(mae[selected_idx]),
            "选择期RMSE": float(rmse[selected_idx]),
            "选择期未触顶MAE": float(noncap_mae[selected_idx]),
            "选择_Logistic_L2": float(chosen["logit_l2"]),
            "选择_Ridge_alpha": float(chosen["ridge_alpha"]),
            "概率定价权重": float(chosen["probability_weight"]),
            "概率校准窗口": int(chosen["prob_cal_window"]),
            "价格校准窗口": int(chosen["price_cal_window"]),
        })
    return pd.DataFrame(records)


def grouped_stats(frame: pd.DataFrame, group_col: str) -> pd.DataFrame:
    rows = []
    for label, part in frame.groupby(group_col, dropna=False, sort=True):
        met = stats(
            part["实际上市收盘价"].to_numpy(float),
            part["自适应预测价"].to_numpy(float),
        )
        rows.append({group_col: label, **met})
    return pd.DataFrame(rows)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--project-dir",
        default=str(PROJECT_DIR),
    )
    args = parser.parse_args()
    review_dir = Path(args.project_dir) / "复核"
    version_dir = Path(args.project_dir) / "版本登记"
    review_dir.mkdir(parents=True, exist_ok=True)
    version_dir.mkdir(parents=True, exist_ok=True)

    model = import_model()
    analysis = model.import_analysis()
    history = model.load_history_with_latest_actual(analysis)
    raw = model.walk_forward_raw(history)
    meta, matrix, baseline_idx = candidate_matrix(model, raw)
    outer = select_outer(model, raw, meta, matrix, baseline_idx)
    if outer.empty:
        raise RuntimeError("没有形成有效的外层逐券预测，请检查热身样本要求")

    outer["年份"] = pd.to_datetime(outer["上市日期"]).dt.year
    outer["触顶状态"] = np.where(outer["实际触顶"].astype(bool), "触及157.30元", "未触顶")
    adaptive = stats(outer["实际上市收盘价"].to_numpy(float), outer["自适应预测价"].to_numpy(float))
    baseline = stats(outer["实际上市收盘价"].to_numpy(float), outer["固定权重基准预测价"].to_numpy(float))

    err = outer["自适应误差"].to_numpy(float)
    ae = np.abs(err)
    signed_quantiles = {
        f"p{int(q * 100):02d}": float(np.quantile(err, q))
        for q in [0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95]
    }
    absolute_quantiles = {
        f"p{int(q * 100):02d}": float(np.quantile(ae, q))
        for q in [0.50, 0.80, 0.90, 0.95, 1.00]
    }
    bin_edges = [-np.inf, -20, -10, -5, 0, 5, 10, 20, np.inf]
    bin_labels = ["<-20", "-20~-10", "-10~-5", "-5~0", "0~5", "5~10", "10~20", ">20"]
    binned = pd.cut(err, bins=bin_edges, labels=bin_labels, right=False)
    counts = pd.Series(binned).value_counts(sort=False)
    error_bins = pd.DataFrame({
        "误差区间": bin_labels,
        "样本数": counts.reindex(bin_labels, fill_value=0).to_numpy(int),
    })
    error_bins["占比"] = error_bins["样本数"] / len(outer)

    by_year = grouped_stats(outer, "年份")
    by_state = grouped_stats(outer, "触顶状态")
    all_valid_raw = np.isfinite(next(iter(raw.prob.values())))
    first_raw_date = pd.to_datetime(raw.frame.loc[all_valid_raw, "上市日期"]).min()
    summary = {
        "model_version": "cb_listing_v2.0_adaptive_outer_walk_forward",
        "rule_start": str(model.RULE_START.date()),
        "history_count_since_rule_start": int(len(history)),
        "first_raw_model_prediction_date": str(first_raw_date.date()),
        "first_outer_adaptive_prediction_date": str(pd.to_datetime(outer["上市日期"]).min().date()),
        "last_prediction_date": str(pd.to_datetime(outer["上市日期"]).max().date()),
        "outer_prediction_count": int(len(outer)),
        "warmup_excluded_count": int(len(history) - len(outer)),
        "adaptive": adaptive,
        "fixed_weight_baseline_same_sample": baseline,
        "mae_improvement_vs_baseline": float(1 - adaptive["mae"] / baseline["mae"]),
        "signed_error_quantiles": signed_quantiles,
        "absolute_error_quantiles": absolute_quantiles,
        "selection_rule": "for each bond, choose hyperparameters only from prior three-year walk-forward errors; same-day listings are excluded from calibration and selection",
    }

    outer.to_csv(review_dir / "v2_outer_predictions.csv", index=False, encoding="utf-8-sig")
    error_bins.to_csv(review_dir / "v2_error_bins.csv", index=False, encoding="utf-8-sig")
    by_year.to_csv(review_dir / "v2_backtest_by_year.csv", index=False, encoding="utf-8-sig")
    by_state.to_csv(review_dir / "v2_backtest_by_cap_state.csv", index=False, encoding="utf-8-sig")
    meta.to_csv(version_dir / "v2_candidate_grid.csv", index=False, encoding="utf-8-sig")
    (review_dir / "v2_backtest_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print("\nBY YEAR\n", by_year.to_string(index=False))
    print("\nBY CAP STATE\n", by_state.to_string(index=False))
    print("\nERROR BINS\n", error_bins.to_string(index=False))


if __name__ == "__main__":
    main()
