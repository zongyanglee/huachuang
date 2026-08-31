from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import unittest

import pandas as pd


SCRIPT = Path(__file__).resolve().parents[1] / "src" / "daily" / "【日报】转债日报.py"
SPEC = importlib.util.spec_from_file_location("cb_daily_industry_history", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class IndustryHistoryAggregationTests(unittest.TestCase):
    def test_lifecycle_demon_filter_and_full_industry_grid_are_preserved(self) -> None:
        """错误纳入未上市券、妖债或删除空行业行时，本测试应失败。"""
        panel = pd.DataFrame(
            {
                "转债代码": ["A", "B", "A", "B", "C", "C"],
                "交易日期": pd.to_datetime(
                    ["2019-01-02", "2019-01-02", "2019-01-03", "2019-01-03", "2019-01-03", "2019-01-03"]
                ),
                "交易状态": ["交易", "交易", "交易", "交易", "交易", "停牌"],
                "收盘价": [100.0, 130.0, 200.0, 120.0, 110.0, 300.0],
                "转股溢价率": [20.0, 10.0, 60.0, 10.0, 15.0, 5.0],
                "平价": [90.0, 100.0, 80.0, 110.0, 105.0, 200.0],
                "纯债溢价率": [30.0, 25.0, 80.0, None, 35.0, 10.0],
            }
        )
        master = pd.DataFrame(
            {
                "转债代码": ["A", "B", "C"],
                "上市日期": pd.to_datetime(["2019-01-01", "2019-01-03", "2019-01-01"]),
                "最后交易日": [pd.NaT, pd.NaT, pd.NaT],
                "申万行业": ["国防军工", "国防军工", "银行"],
            }
        )

        self.assertTrue(
            hasattr(MODULE, "aggregate_industry_history_metrics"),
            "日报模块尚未提供行业历史序列聚合函数",
        )
        result = MODULE.aggregate_industry_history_metrics(
            panel,
            master,
            start_date=pd.Timestamp("2019-01-01"),
            run_date=pd.Timestamp("2019-01-03"),
        )

        self.assertEqual(len(result), 60)
        self.assertEqual(result["行业"].drop_duplicates().tolist(), list(MODULE.INDUSTRY_HISTORY_ORDER))
        self.assertEqual(
            result.loc[result["行业"].eq("国防军工"), "日期"].dt.strftime("%Y-%m-%d").tolist(),
            ["2019-01-02", "2019-01-03"],
        )
        defense = result.loc[result["行业"].eq("国防军工")].reset_index(drop=True)
        self.assertEqual(defense.loc[0, "收盘价"], 100.0)
        self.assertEqual(defense.loc[1, "收盘价"], 120.0)
        self.assertTrue(pd.isna(defense.loc[1, "纯债溢价率"]))
        bank = result.loc[result["行业"].eq("银行")].reset_index(drop=True)
        self.assertEqual(bank.loc[1, "收盘价"], 110.0)
        self.assertTrue(result.loc[result["行业"].eq("房地产"), "收盘价"].isna().all())


if __name__ == "__main__":
    unittest.main()
