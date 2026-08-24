import json
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
INPUT_PATH = ROOT / "0710数据更新" / "日内估值数据更新" / "0710日内数据更新.xlsx"
OUTPUT_DIR = ROOT / "outputs" / "parity_group_premium_stats_20260710"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

GROUPS = [
    ("70-90", 70.0, 90.0),
    ("90-110", 90.0, 110.0),
    ("110-130", 110.0, 130.0),
    ("130-150", 130.0, 150.0),
]


def _wide_from_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name)
    first_col = df.columns[0]
    return df.rename(columns={first_col: "转债代码"}).set_index("转债代码")


def _weighted_mean(values: pd.Series, weights: pd.Series) -> float:
    valid = values.notna() & weights.notna() & (weights > 0)
    if not valid.any():
        return np.nan
    return float(np.average(values[valid], weights=weights[valid]))


def main() -> None:
    parity = _wide_from_sheet(INPUT_PATH, "平价")
    premium = _wide_from_sheet(INPUT_PATH, "转股溢价率")
    balance = _wide_from_sheet(INPUT_PATH, "转债余额")

    balance_col = balance.columns[0]
    balance_series = pd.to_numeric(balance[balance_col], errors="coerce")

    common_codes = parity.index.intersection(premium.index).intersection(balance_series.index)
    common_cols = [col for col in parity.columns if col in premium.columns]

    rows = []
    for timestamp in common_cols:
        parity_values = pd.to_numeric(parity.loc[common_codes, timestamp], errors="coerce")
        premium_values = pd.to_numeric(premium.loc[common_codes, timestamp], errors="coerce")
        balances = balance_series.loc[common_codes]

        for label, lower, upper in GROUPS:
            mask = (parity_values > lower) & (parity_values <= upper) & premium_values.notna()
            sub_premium = premium_values[mask]
            sub_balance = balances[mask]
            rows.append(
                {
                    "timestamp": str(timestamp),
                    "group": label,
                    "lower_open": lower,
                    "upper_closed": upper,
                    "sample_count": int(sub_premium.shape[0]),
                    "balance_sum": float(sub_balance.dropna().sum()),
                    "arithmetic_mean_premium": float(sub_premium.mean()) if not sub_premium.empty else None,
                    "balance_weighted_mean_premium": _weighted_mean(sub_premium, sub_balance),
                    "median_premium": float(sub_premium.median()) if not sub_premium.empty else None,
                }
            )

    stats = pd.DataFrame(rows)
    latest_timestamp = common_cols[-1] if common_cols else None
    latest = stats[stats["timestamp"].eq(str(latest_timestamp))].copy()

    payload = {
        "source_file": str(INPUT_PATH.relative_to(ROOT)),
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        "latest_timestamp": str(latest_timestamp),
        "group_rule": "按平价分组，区间为左开右闭：(70,90]、(90,110]、(110,130]、(130,150]。",
        "metric_rule": "算术平均和中位数使用转股溢价率有效样本；余额加权平均使用余额>0且转股溢价率有效样本，公式为 sum(转股溢价率*余额)/sum(余额)。转股溢价率单位为百分点。",
        "groups": [label for label, _, _ in GROUPS],
        "latest_rows": latest.replace({np.nan: None}).to_dict(orient="records"),
        "time_series_rows": stats.replace({np.nan: None}).to_dict(orient="records"),
    }

    json_path = OUTPUT_DIR / "parity_group_premium_stats.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json_path)
    print("latest_timestamp", latest_timestamp)
    print(latest[["group", "sample_count", "arithmetic_mean_premium", "balance_weighted_mean_premium", "median_premium"]].to_string(index=False))


if __name__ == "__main__":
    main()
