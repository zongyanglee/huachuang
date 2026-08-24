from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List

import numpy as np
import pandas as pd

sys.path.append(str(Path(__file__).resolve().parents[1]))

from 底稿更新 import load_original_data


@dataclass(frozen=True)
class RunConfig:
    parquet_root: str = "data/转债个券历史序列"
    output_xlsx: str = "银行转债_余额及个数_时间序列_parquet.xlsx"

    total_sheet: str = "总表"
    industry_col: str = "申万行业"
    industry_value: str = "银行"

    balance_sheet: str = "余额"


def _date_cols(df: pd.DataFrame) -> List[pd.Timestamp]:
    return sorted([c for c in df.columns if isinstance(c, pd.Timestamp)])


def _as_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def main(config: RunConfig) -> Path:
    data = load_original_data(source_type="parquet", parquet_root=config.parquet_root, force_refresh=False)
    for sheet in [config.total_sheet, config.balance_sheet]:
        if sheet not in data:
            raise KeyError(f"parquet 数据缺少 sheet: {sheet}")

    total_df = data[config.total_sheet].copy()
    if "__row_id" in total_df.columns:
        total_df = total_df.set_index("__row_id")

    industry_s = total_df.get(config.industry_col)
    if industry_s is None:
        raise KeyError(f"`{config.total_sheet}` 中未找到 `{config.industry_col}` 字段。")

    bal_df = data[config.balance_sheet]
    dates = _date_cols(bal_df)
    if not dates:
        raise ValueError(f"`{config.balance_sheet}` 未找到日期列。")

    # 银行转债池（静态口径：总表.申万行业 == 银行）
    bank_mask = industry_s.reindex(bal_df.index).astype("object").eq(config.industry_value)
    bank_codes = bank_mask[bank_mask].index

    rows = []
    for d in dates:
        bal = _as_num(bal_df.loc[bank_codes, d])
        ok = bal.notna() & np.isfinite(bal) & (bal > 0)
        rows.append(
            {
                "日期": d,
                "银行转债个数": int(ok.sum()),
                "银行转债余额合计": float(bal[ok].sum()) if ok.any() else float("nan"),
            }
        )

    out_df = pd.DataFrame(rows).set_index("日期").sort_index()

    out_path = Path(config.output_xlsx).resolve()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        info = pd.DataFrame(
            [
                {"key": "数据源", "value": f"parquet: {Path(config.parquet_root).resolve()}"},
                {"key": "行业口径", "value": f"{config.total_sheet}.{config.industry_col} == {config.industry_value}"},
                {"key": "余额口径", "value": "同日余额>0 计入个数与合计"},
                {"key": "日期范围", "value": f"{out_df.index.min().date()} ~ {out_df.index.max().date()} (n={len(out_df)})"},
                {"key": "银行转债池规模(静态)", "value": int(len(bank_codes))},
            ]
        )
        info.to_excel(writer, sheet_name="说明", index=False)
        out_df.to_excel(writer, sheet_name="时间序列", index=True)

    return out_path


if __name__ == "__main__":
    out = main(RunConfig())
    print(f"[ok] wrote: {out}")
