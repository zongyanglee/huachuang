from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import unittest

import pandas as pd


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_parity_interval", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailyParityIntervalPremiumTests(unittest.TestCase):
    def test_uses_daily_update_buckets_and_balance_weighting(self) -> None:
        """七档边界及余额加权结果应与日度数据更新口径一致。"""
        rows = [
            ("2026-08-28", "交易", 1.0, 75.0, 10.0),
            ("2026-08-28", "交易", 3.0, 80.0, 30.0),
            ("2026-08-28", "交易", 2.0, 85.0, 20.0),
            ("2026-08-28", "交易", 2.0, 90.0, 40.0),
            ("2026-08-28", "交易", 1.0, 95.0, 50.0),
            ("2026-08-28", "交易", 3.0, 100.0, 10.0),
            ("2026-08-28", "交易", 1.0, 105.0, 20.0),
            ("2026-08-28", "交易", 3.0, 110.0, 40.0),
            ("2026-08-28", "交易", 1.0, 115.0, 30.0),
            ("2026-08-28", "交易", 3.0, 120.0, 50.0),
            ("2026-08-28", "交易", 1.0, 125.0, 40.0),
            ("2026-08-28", "交易", 3.0, 130.0, 60.0),
            ("2026-08-28", "交易", 2.0, 135.0, 70.0),
            ("2026-08-28", "未上市", 100.0, 75.0, 999.0),
            ("2026-08-29", "交易", 2.0, 135.0, 80.0),
        ]
        data = pd.DataFrame(
            rows,
            columns=["交易日期", "交易状态", "余额", "平价", "转股溢价率"],
        )
        aggregate = getattr(MODULE, "aggregate_parity_interval_premium_series", None)
        self.assertIsNotNone(aggregate)

        result = aggregate(data)

        self.assertEqual(
            result.columns.tolist(),
            [
                "交易日期",
                "130以上",
                "120-130（含130）",
                "110-120（含120）",
                "100-110（含110）",
                "90-100（含100）",
                "80-90（含90）",
                "80以下（含80）",
            ],
        )
        latest = result.loc[result["交易日期"].eq(pd.Timestamp("2026-08-28"))].iloc[0]
        self.assertEqual(float(latest["130以上"]), 70.0)
        self.assertEqual(float(latest["120-130（含130）"]), 55.0)
        self.assertEqual(float(latest["110-120（含120）"]), 45.0)
        self.assertEqual(float(latest["100-110（含110）"]), 35.0)
        self.assertEqual(float(latest["90-100（含100）"]), 20.0)
        self.assertEqual(float(latest["80-90（含90）"]), 30.0)
        self.assertEqual(float(latest["80以下（含80）"]), 25.0)


if __name__ == "__main__":
    unittest.main()
