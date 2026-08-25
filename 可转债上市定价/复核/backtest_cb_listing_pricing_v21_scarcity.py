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
V21_PATH = PROJECT_DIR / "版本登记" / "cb_listing_pricing_v21_scarcity.py"
V2_BACKTEST_PATH = PROJECT_DIR / "复核" / "backtest_cb_listing_pricing_v2_adaptive.py"


def import_file(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--project-dir",
        default=str(PROJECT_DIR),
    )
    args = parser.parse_args()
    review_dir = Path(args.project_dir) / "复核"
    review_dir.mkdir(parents=True, exist_ok=True)

    v21 = import_file("cb_listing_pricing_v21", V21_PATH)
    bt = import_file("cb_listing_pricing_v2_bt", V2_BACKTEST_PATH)
    v2 = v21.import_file("cb_listing_pricing_v2_for_v21_bt", v21.V2_PATH)
    analysis = v2.import_analysis()
    history = v2.load_history_with_latest_actual(analysis)
    raw = v2.walk_forward_raw(history)
    meta, matrix, baseline_idx = bt.candidate_matrix(v2, raw)
    outer_v2 = bt.select_outer(v2, raw, meta, matrix, baseline_idx)

    rows = []
    for _, row in outer_v2.iterrows():
        target_date = pd.Timestamp(row["上市日期"])
        match = (
            meta["logit_l2"].eq(float(row["选择_Logistic_L2"]))
            & meta["ridge_alpha"].eq(float(row["选择_Ridge_alpha"]))
            & meta["probability_weight"].eq(float(row["概率定价权重"]))
            & meta["prob_cal_window"].eq(int(row["概率校准窗口"]))
            & meta["price_cal_window"].eq(int(row["价格校准窗口"]))
        )
        candidate_idx = int(np.flatnonzero(match.to_numpy())[0])
        base_history = matrix[candidate_idx].astype(float)
        scarcity, _ = v21.select_scarcity_overlay(raw.frame, base_history, target_date)
        target_chips = float(
            row["发行规模亿元"] * (1.0 - row["原股东配售率"])
        )
        correction = float(v21.scarcity_adjustment(
            target_chips,
            scarcity["筹码阈值亿元"],
            scarcity["筹码斜率"],
            scarcity["单券修正上限"],
        ))
        final = float(min(v2.CAP, row["自适应预测价"] + correction))
        rows.append({
            **row.to_dict(),
            "模型版本": "v2.1_极低筹码非线性修正",
            "潜在首日流通筹码亿元": target_chips,
            **scarcity,
            "v2.0预测价": float(row["自适应预测价"]),
            "筹码非线性修正额": correction,
            "v2.1预测价": final,
            "v2.1误差": final - float(row["实际上市收盘价"]),
            "v2.1绝对误差": abs(final - float(row["实际上市收盘价"])),
        })

    result = pd.DataFrame(rows)
    actual = result["实际上市收盘价"].to_numpy(float)
    pred_v2 = result["v2.0预测价"].to_numpy(float)
    pred_v21 = result["v2.1预测价"].to_numpy(float)
    hit = actual >= 157.29
    old = v21.error_metrics(actual, pred_v2, hit)
    new = v21.error_metrics(actual, pred_v21, hit)
    result["年份"] = pd.to_datetime(result["上市日期"]).dt.year
    by_year = []
    for year, part in result.groupby("年份"):
        a = part["实际上市收盘价"].to_numpy(float)
        h = a >= 157.29
        by_year.append({
            "年份": int(year),
            **{f"v2.0_{k}": v for k, v in v21.error_metrics(a, part["v2.0预测价"].to_numpy(float), h).items()},
            **{f"v2.1_{k}": v for k, v in v21.error_metrics(a, part["v2.1预测价"].to_numpy(float), h).items()},
        })
    by_year = pd.DataFrame(by_year)
    tail = result.loc[result["v2.1误差"].lt(-10)].copy()
    large_error = result.loc[result["v2.1误差"].abs().gt(10)].copy()
    large_error["误差方向"] = np.where(large_error["v2.1误差"].lt(0), "低估", "高估")

    summary = {
        "model_version": "cb_listing_v2.1_scarcity_outer_walk_forward",
        "scarcity_factor": "发行规模亿元 × (1 - 原股东配售率)",
        "scarcity_formula": "min(单券修正上限, 筹码斜率 × max(0, ln(筹码阈值/潜在首日流通筹码)))",
        "parameter_grids": {
            "threshold_billion_yuan": v21.CHIP_THRESHOLD_GRID,
            "slope": v21.CHIP_SLOPE_GRID,
            "correction_cap_yuan": v21.CHIP_CAP_GRID,
        },
        "selection_constraints": "each bond uses only prior-three-year errors; RMSE, overprediction-over-10 count and non-cap MAE cannot be worse than v2.0 in the selection window",
        "first_prediction_date": str(pd.to_datetime(result["上市日期"]).min().date()),
        "last_prediction_date": str(pd.to_datetime(result["上市日期"]).max().date()),
        "n": int(len(result)),
        "v2.0": old,
        "v2.1": new,
        "mae_improvement": float(1 - new["mae"] / old["mae"]),
        "rmse_improvement": float(1 - new["rmse"] / old["rmse"]),
        "corrected_bond_count": int(np.sum(result["筹码非线性修正额"].gt(0))),
        "underestimation_over_10_count": int(len(tail)),
        "absolute_error_over_10_count": int(len(large_error)),
    }
    result.to_csv(review_dir / "v21_outer_predictions.csv", index=False, encoding="utf-8-sig")
    by_year.to_csv(review_dir / "v21_backtest_by_year.csv", index=False, encoding="utf-8-sig")
    tail.to_csv(review_dir / "v21_underestimation_over_10.csv", index=False, encoding="utf-8-sig")
    large_error.to_csv(review_dir / "v21_absolute_error_over_10.csv", index=False, encoding="utf-8-sig")
    (review_dir / "v21_backtest_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print("\nBY YEAR\n", by_year.to_string(index=False))


if __name__ == "__main__":
    main()
