from __future__ import annotations

import importlib.util
from datetime import date
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
SPEC = importlib.util.spec_from_file_location("cb_daily_commentary", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class DailyCommentaryTests(unittest.TestCase):
    def test_price_aggregation_includes_balance_weighted_close(self) -> None:
        """价格聚合应补充全样本余额加权收盘价，供点评文字使用。"""
        data = pd.DataFrame(
            {
                "交易日期": ["2026-08-27"] * 3 + ["2026-08-28"] * 2,
                "交易状态": ["交易", "交易", "未上市", "交易", "交易"],
                "余额": [1.0, 3.0, 100.0, 2.0, 2.0],
                "平价": [90.0, 110.0, 999.0, 100.0, 120.0],
                "收盘价": [100.0, 200.0, 999.0, 120.0, 180.0],
            }
        )
        aggregate = getattr(MODULE, "aggregate_price_parity_series", None)
        self.assertIsNotNone(aggregate)

        result = aggregate(data)

        first = result.loc[result["交易日期"].eq(pd.Timestamp("2026-08-27"))].iloc[0]
        self.assertEqual(float(first["余额加权平价"]), 105.0)
        self.assertEqual(float(first["余额加权收盘价"]), 175.0)
        self.assertEqual(float(first["收盘价中位数"]), 150.0)

    def test_commentary_uses_requested_sections_and_dynamic_values(self) -> None:
        """点评应按指定段落顺序输出动态市场、估值和行业数据。"""
        index_changes = {
            "中证转债": 0.20,
            "上证综指": -0.10,
            "深证成指": -0.30,
            "创业板指": 0.40,
            "上证50": 0.00,
            "中证1000": 0.50,
            "大盘成长": -0.60,
            "大盘价值": -0.20,
            "中盘成长": -0.40,
            "中盘价值": 0.30,
            "小盘成长": -0.50,
            "小盘价值": 0.10,
        }
        index_performance = pd.DataFrame(
            [
                {"指数名称": name, "日涨跌幅": change}
                for name, change in index_changes.items()
            ]
        )
        turnover = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2026-08-27", "2026-08-28"]),
                "中证转债指数成交额_亿元": [500.0, 450.0],
                "沪深成交额合计_亿元": [10000.0, 11000.0],
            }
        )
        price_parity_source = {
            "latestDate": "2026-08-28",
            "previousDate": "2026-08-27",
            "latestWeightedClose": 120.0,
            "weightedCloseDailyChangePct": 20.0,
            "latestParity": 95.0,
            "parityDailyChangePct": 2.0,
            "latestMedianPrice": 110.0,
            "medianPriceDailyChangePct": -1.0,
        }
        equity_bond_weighted_source = {
            "price": {
                "偏股型": {"latestValue": 200.0, "dailyChangePct": 2.0},
                "平衡型": {"latestValue": 130.0, "dailyChangePct": 1.0},
                "偏债型": {"latestValue": 110.0, "dailyChangePct": -1.0},
            }
        }
        close_distribution = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2026-08-27", "2026-08-28"]),
                "150以上": [25.0, 20.0],
            }
        )
        close_distribution_source = {
            "latestBreakFloorPct": 2.5,
            "breakFloorDailyChangePctPoint": 0.3,
        }
        valuation_source = {
            "latestValue": 40.0,
            "dailyChangePctPoint": -1.0,
            "previousDate": "2026-08-27",
        }
        valuation_repair = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2026-08-27", "2026-08-28"]),
                MODULE.VALUATION_REPAIR_INDEX_NAME: [20.0, 18.0],
            }
        )
        subnew_source = {
            "latestPremiumMeanPct": 50.0,
            "premiumDailyChangePctPoint": 1.0,
        }
        fitted_details = {
            "偏股型": {"latestValue": 30.0, "dailyChangePctPoint": -1.0},
            "平衡型": {"latestValue": 40.0, "dailyChangePctPoint": 2.0},
            "偏债型": {"latestValue": 80.0, "dailyChangePctPoint": -3.0},
        }
        parity_details = {
            "70-90": {"latestValue": 60.0, "dailyChangePctPoint": -1.0},
            "90-110": {"latestValue": 40.0, "dailyChangePctPoint": 0.5},
            "110-130": {"latestValue": 30.0, "dailyChangePctPoint": -0.5},
        }
        industries = pd.DataFrame(
            {
                "行业名称": ["行业甲", "行业乙", "行业丙", "行业丁", "行业戊", "行业己"],
                "正股日涨跌幅": [3.0, 2.0, 1.0, -1.0, -2.0, -3.0],
                "转债日涨跌幅": [1.0, 0.5, -3.0, -2.0, 2.0, -1.0],
            }
        )
        sector_metrics = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2026-08-27", "2026-08-28"]),
                **{
                    f"{metric}_{sector}": values
                    for metric, values in {
                        "收盘价": [100.0, 101.0],
                        "平价": [100.0, 102.0],
                        "转股溢价率": [40.0, 39.0],
                        "纯债溢价率": [20.0, 19.5],
                    }.items()
                    for sector in MODULE.SECTOR_ORDER
                },
            }
        )
        builder = getattr(MODULE, "build_daily_commentary", None)
        self.assertIsNotNone(builder)

        text = builder(
            date(2026, 8, 28),
            index_performance=index_performance,
            turnover=turnover,
            main_money_flow_source={"latestValue": -12.34},
            etf_share_source={
                "funds": {
                    "博时可转债ETF": {
                        "latestShareYi": 45.5,
                        "latestNetSubscriptionYi": -0.000211,
                    },
                    "海富通可转债ETF": {
                        "latestShareYi": 9.88,
                        "latestNetSubscriptionYi": 0.02,
                    },
                }
            },
            price_parity_source=price_parity_source,
            equity_bond_weighted_source=equity_bond_weighted_source,
            close_price_distribution=close_distribution,
            close_price_distribution_source=close_distribution_source,
            valuation_source=valuation_source,
            valuation_repair=valuation_repair,
            subnew_bond_source=subnew_source,
            equity_bond_group_valuation_source={"groupDetails": fitted_details},
            parity_group_valuation_source={"groupDetails": parity_details},
            industry_performance=industries,
            sector_mean_metrics=sector_metrics,
        )

        expected_sections = [
            "市场概况：",
            "指数表现：",
            "市场风格：",
            "资金表现：",
            "转债价格：",
            "转债估值：",
            "行业表现：",
            "(1) 收盘价：",
            "(2) 转股溢价率：",
            "(3) 转换价值：",
            "(4) 纯债溢价率：",
        ]
        positions = [text.index(section) for section in expected_sections]
        self.assertEqual(positions, sorted(positions))
        self.assertIn("市场概况：转债缩量上涨，估值压缩", text)
        self.assertIn("中盘价值相对占优", text)
        self.assertIn("可转债市场成交额为450.00亿元，环比减少10.00%", text)
        self.assertIn("博时可转债ETF份额45.50亿份，净申赎-2.11万份", text)
        self.assertIn("转债整体收盘价加权平均值为120.00元", text)
        self.assertIn("破底占比为2.50%，环比+0.30pct", text)
        self.assertIn("转债估值修复指数18.00%；-2.00pct", text)
        self.assertIn("涨幅前三位行业为行业甲（+3.00%）", text)
        self.assertIn("大周期环比+1.00%", text)
        self.assertNotIn("xx.xx", text)

    def test_long_chart_titles_are_appended_in_report_reading_order(self) -> None:
        """点评末尾应追加长图全部24个图表标题，并保留标题内部换行。"""
        builder = getattr(MODULE, "build_long_chart_titles", None)
        appender = getattr(MODULE, "append_long_chart_titles", None)
        self.assertIsNotNone(builder)
        self.assertIsNotNone(appender)
        turnover = pd.DataFrame(
            {
                "交易日期": pd.to_datetime(["2026-08-27", "2026-08-28"]),
                "中证转债指数成交额_亿元": [400.0, 500.0],
                "沪深成交额合计_亿元": [10000.0, 11000.0],
            }
        )
        group_titles = {
            "parity": "平价分类标题",
            "equity": "股债型标题",
            "rating": "评级标题",
            "balance": "余额标题",
            "market_cap": "市值标题",
            "sector": "板块标题",
        }
        titles = builder(
            turnover=turnover,
            return_summary={"上涨": 60, "下跌": 40, "有效样本": 100},
            main_money_flow_source={"latestValue": -12.34},
            margin_balance=pd.DataFrame(
                {
                    "交易日期": pd.to_datetime(["2026-08-27", "2026-08-28"]),
                    "沪深两市融资融券余额_亿元": [20000.0, 20100.0],
                }
            ),
            valuation_source={
                "latestValue": 40.0,
                "dailyChangePctPoint": -1.0,
                "percentileSince2019": 90.0,
            },
            valuation_repair=pd.DataFrame(
                {
                    "交易日期": pd.to_datetime(["2026-08-27", "2026-08-28"]),
                    MODULE.VALUATION_REPAIR_INDEX_NAME: [10.0, 12.0],
                }
            ),
            price_parity_source={
                "latestParity": 95.0,
                "parityDailyChangePct": 1.0,
                "latestMedianPrice": 120.0,
                "medianPriceDailyChangePct": -1.0,
                "medianPricePercentileSince2019": 80.0,
            },
            equity_bond_weighted_source={
                "premium": {
                    "偏股型": {"latestValue": 30.0, "dailyChangePctPoint": -1.0},
                    "偏债型": {"latestValue": 80.0, "dailyChangePctPoint": 2.0},
                },
                "price": {
                    "偏股型": {"latestValue": 200.0, "dailyChangePct": 1.0},
                    "偏债型": {"latestValue": 110.0, "dailyChangePct": -1.0},
                },
            },
            group_valuation_titles=group_titles,
            maturity_title="剩余期限标题",
            subnew_bond_source={
                "latestPremiumMeanPct": 50.0,
                "premiumDailyChangePctPoint": 1.0,
            },
            close_price_distribution_source={
                "latestBreakFloorPct": 2.0,
                "breakFloorDailyChangePctPoint": 0.1,
                "latestBreakParPct": 10.0,
                "breakParDailyChangePctPoint": -0.2,
            },
            etf_share_source={
                "funds": {
                    "博时可转债ETF": {
                        "latestShareYi": 45.5,
                        "latestNetSubscriptionYi": -0.0002,
                    },
                    "海富通可转债ETF": {
                        "latestShareYi": 9.8,
                        "latestNetSubscriptionYi": 0.02,
                    },
                }
            },
        )

        self.assertEqual(len(titles), 24)
        self.assertTrue(titles[0].startswith("成交额:转债500.00亿"))
        self.assertTrue(titles[1].startswith("上涨转债占比60.00%"))
        self.assertEqual(titles[6], "转债估值修复指数12.00%；+2.00pct")
        self.assertEqual(titles[9], "平价分类标题")
        self.assertEqual(titles[20:], [item[1] for item in MODULE.SECTOR_MEAN_METRICS])
        output = appender("原点评", titles)
        self.assertIn("原点评\n\n图表标题：\n", output)
        self.assertIn("百元拟合溢价率：40.00%，-1.00pct\n2019年以来90.00%分位数", output)


if __name__ == "__main__":
    unittest.main()
