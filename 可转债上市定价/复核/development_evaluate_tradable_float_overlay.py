from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
PROJECT_DIR = Path(__file__).resolve().parents[1]
SOURCE = PROJECT_DIR / "复核" / "v2_outer_predictions.csv"


def metrics(actual, pred):
    err = pred - actual
    return {
        "mae": float(np.mean(np.abs(err))),
        "rmse": float(np.sqrt(np.mean(err**2))),
        "within5": float(np.mean(np.abs(err) <= 5)),
        "under10": int(np.sum(err < -10)),
        "over10": int(np.sum(err > 10)),
        "bias": float(np.mean(err)),
    }


data = pd.read_csv(SOURCE, parse_dates=["上市日期"]).sort_values(
    ["上市日期", "转债代码"]
).reset_index(drop=True)
data["可交易筹码亿元"] = data["发行规模亿元"] * (1 - data["原股东配售率"])
thresholds = [0.50, 0.75, 1.00, 1.25, 1.50, 2.00]
premiums = np.arange(0, 12.01, 1.0)
records = []

for i, row in data.iterrows():
    date = row["上市日期"]
    prior = data.loc[
        data["上市日期"].lt(date)
        & data["上市日期"].ge(date - pd.DateOffset(years=3))
    ]
    if len(prior) < 30:
        continue
    actual = prior["实际上市收盘价"].to_numpy(float)
    base = prior["自适应预测价"].to_numpy(float)
    chips = prior["可交易筹码亿元"].to_numpy(float)
    baseline = metrics(actual, base)
    candidates = []
    for threshold in thresholds:
        for premium in premiums:
            pred = np.minimum(157.30, base + (chips < threshold) * premium)
            met = metrics(actual, pred)
            if met["rmse"] <= baseline["rmse"] + 1e-12 and met["over10"] <= baseline["over10"]:
                candidates.append((met["mae"], met["rmse"], premium, threshold, met))
    selected = min(candidates) if candidates else (baseline["mae"], baseline["rmse"], 0.0, 1.0, baseline)
    _, _, premium, threshold, selection_metrics = selected
    corrected = min(
        157.30,
        float(row["自适应预测价"]) + (premium if float(row["可交易筹码亿元"]) < threshold else 0.0),
    )
    records.append({
        "转债代码": row["转债代码"],
        "转债名称": row["转债名称"],
        "上市日期": date,
        "实际价": float(row["实际上市收盘价"]),
        "原预测价": float(row["自适应预测价"]),
        "修正预测价": corrected,
        "可交易筹码亿元": float(row["可交易筹码亿元"]),
        "选择阈值亿元": threshold,
        "选择加价元": premium,
        "选择样本数": len(prior),
    })

result = pd.DataFrame(records)
base_metrics = metrics(result["实际价"].to_numpy(float), result["原预测价"].to_numpy(float))
overlay_metrics = metrics(result["实际价"].to_numpy(float), result["修正预测价"].to_numpy(float))
print("DATE", result["上市日期"].min().date(), result["上市日期"].max().date(), "N", len(result))
print("BASE", base_metrics)
print("OVERLAY", overlay_metrics)
print("MAE_IMPROVEMENT", 1 - overlay_metrics["mae"] / base_metrics["mae"])
print("\nPARAMETERS")
print(
    result.groupby(["选择阈值亿元", "选择加价元"]).size().sort_values(ascending=False).head(12).to_string()
)
print("\nCORRECTED COUNT", int((result["修正预测价"] != result["原预测价"]).sum()))
