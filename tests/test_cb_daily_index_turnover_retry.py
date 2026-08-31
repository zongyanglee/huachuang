from __future__ import annotations

from contextlib import redirect_stderr
from datetime import date
import importlib.util
import io
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import pandas as pd


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_report", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class FakeResult:
    def __init__(
        self,
        errorcode: int,
        errmsg: str,
        data: pd.DataFrame | None = None,
    ) -> None:
        self.errorcode = errorcode
        self.errmsg = errmsg
        self.data = data


class IndexTurnoverRetryTests(unittest.TestCase):
    def test_transient_request_error_is_retried_before_succeeding(self) -> None:
        successful_data = pd.DataFrame(
            {
                "time": ["2026-08-28"],
                "000001.SH": [1_000_000_000],
                "399001.SZ": [2_000_000_000],
                "000832.CSI": [300_000_000],
            }
        )
        responses = [
            FakeResult(-205, "request data error"),
            FakeResult(0, "Success!", successful_data),
        ]

        with (
            patch.object(MODULE, "THS_DS", side_effect=responses) as ths_ds,
            patch.object(MODULE.time, "sleep") as sleep,
        ):
            try:
                result = MODULE.fetch_index_turnover(
                    date(2025, 8, 28), date(2026, 8, 28)
                )
            except RuntimeError as exc:
                self.fail(f"瞬时请求错误未被重试：{exc}")

        self.assertEqual(ths_ds.call_count, 2)
        sleep.assert_called_once()
        self.assertEqual(result["交易日期"].dt.strftime("%Y-%m-%d").tolist(), ["2026-08-28"])
        self.assertEqual(result["沪深成交额合计_亿元"].tolist(), [30.0])


class ConsoleProgressTests(unittest.TestCase):
    def test_shorter_chinese_message_clears_previous_terminal_text(self) -> None:
        output = io.StringIO()

        with redirect_stderr(output):
            progress = MODULE.ConsoleProgress(width=4)
            progress.update(16, "读取市场表现")
            progress.fail()

        final_line = output.getvalue().split("\r")[-1].rstrip("\n")
        trailing_spaces = len(final_line) - len(final_line.rstrip(" "))
        self.assertEqual(trailing_spaces, 8)


if __name__ == "__main__":
    unittest.main()
