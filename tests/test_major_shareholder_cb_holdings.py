from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import unittest

import pandas as pd


SCRIPT = Path(__file__).resolve().parents[1] / "src" / "clauses" / "【条款】大股东持债情况.py"
SPEC = importlib.util.spec_from_file_location("major_shareholder_cb_holdings", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class MatchMajorShareholderTests(unittest.TestCase):
    def test_exact_match_returns_rank_and_ratio(self):
        result = MODULE.match_major_shareholder(
            "测试集团有限公司",
            ["甲", "测试集团有限公司", "乙"],
            [1.0, 35.25, 2.0],
        )
        self.assertEqual(result.rank, 2)
        self.assertEqual(result.ratio, 35.25)
        self.assertEqual(result.status, "匹配成功")

    def test_match_ignores_surrounding_whitespace_only(self):
        result = MODULE.match_major_shareholder(" 测试集团 ", ["测试集团"], [10])
        self.assertEqual(result.rank, 1)
        self.assertEqual(result.ratio, 10.0)

    def test_no_fuzzy_match_for_different_legal_names(self):
        result = MODULE.match_major_shareholder("测试集团", ["测试集团有限公司"], [10])
        self.assertIsNone(result.rank)
        self.assertIsNone(result.ratio)
        self.assertEqual(result.status, "前十大中未找到")

    def test_missing_ratio_is_reported(self):
        result = MODULE.match_major_shareholder("测试集团", ["测试集团"], [None])
        self.assertEqual(result.rank, 1)
        self.assertIsNone(result.ratio)
        self.assertEqual(result.status, "名称匹配但持债比例缺失")

    def test_holding_status_contains_rank_and_ratio(self):
        match = MODULE.HolderMatch(rank=3, ratio=6.25, status="匹配成功")
        self.assertEqual(MODULE.format_holding_status(match), "转债第3名，持债6.25%")


class MaturityFilterTests(unittest.TestCase):
    def test_filter_includes_both_boundaries(self):
        basic = pd.DataFrame(
            {"转债简称": ["A", "B", "C", "D"]},
            index=["A.SH", "B.SH", "C.SH", "D.SH"],
        )
        maturity = pd.Series(
            [4.999, 5.0, 5.5, 5.501],
            index=basic.index,
        )
        result = MODULE.filter_maturity_range(basic, maturity)
        self.assertEqual(result.index.tolist(), ["B.SH", "C.SH"])


if __name__ == "__main__":
    unittest.main()
