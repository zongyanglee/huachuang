from __future__ import annotations

import importlib.util
from datetime import date
import hashlib
import io
from pathlib import Path
import re
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

import pandas as pd
from PIL import Image


SCRIPT = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "daily"
    / "【日报】转债日报.py"
)
SPEC = importlib.util.spec_from_file_location("cb_daily_word_template", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SEQUENCE_PARAGRAPH_XML = f"""
<w:p xmlns:w="{W_NS}">
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText> SEQ 图表 \\* ARABIC </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>2</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:rPr><w:b/></w:rPr><w:t>旧标题</w:t></w:r>
</w:p>
"""


def _index_frame() -> pd.DataFrame:
    records: list[dict[str, object]] = []
    for group, specs in (
        ("主要指数", MODULE.MAIN_INDEX_SPECS),
        ("风格指数", MODULE.STYLE_INDEX_SPECS),
    ):
        for offset, (_, _, display_name) in enumerate(specs):
            records.append(
                {
                    "组别": group,
                    "指数名称": display_name,
                    "收盘价": 100 + offset,
                    "日涨跌幅": -0.125 + offset,
                    "近一周涨跌幅": 1 + offset,
                    "近一月涨跌幅": 2 + offset,
                    "年初至今涨跌幅": 3 + offset,
                }
            )
    return pd.DataFrame(records)


def _industry_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "行业名称": ["煤炭", "传媒", "计算机", "电子"],
            "正股日涨跌幅": [2.1, 3.8, 2.2, -1.0],
        }
    )


def _package_hashes(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {
            name: hashlib.sha256(package.read(name)).hexdigest()
            for name in package.namelist()
        }


class DailyWordTemplateTests(unittest.TestCase):
    def test_build_word_index_table_rows_formats_both_groups(self) -> None:
        """图表1应输出模板简称，并按左右两组固定为两位小数。"""
        rows = MODULE.build_word_index_table_rows(_index_frame())

        self.assertEqual(len(rows), 9)
        self.assertEqual(
            rows[0],
            [
                "中证转债",
                "100.00",
                "-0.12",
                "1.00",
                "2.00",
                "3.00",
                "大盘指数",
                "100.00",
                "-0.12",
                "1.00",
                "2.00",
                "3.00",
            ],
        )
        self.assertEqual(rows[1][0], "转债等权")
        self.assertEqual(rows[2][0], "正股等权")
        self.assertEqual(rows[1][6], "中盘指数")

    def test_build_industry_rotation_title_uses_top_three_stock_returns(self) -> None:
        """图表26标题应取正股日涨跌幅前三行业。"""
        self.assertEqual(
            MODULE.build_industry_rotation_title(_industry_frame()),
            "行业轮动情况：传媒、计算机、煤炭领涨",
        )

    def test_template_contract_matches_frozen_reference(self) -> None:
        """模板改版时必须在写入前被结构指纹拦截。"""
        contract = MODULE.inspect_daily_word_template(
            MODULE.DAILY_WORD_TEMPLATE_PATH
        )

        self.assertEqual(contract["sha256"], MODULE.DAILY_WORD_TEMPLATE_SHA256)
        self.assertEqual(contract["topLevelTableCount"], 18)
        self.assertEqual(
            contract["chartImageRelationshipIds"],
            [f"rId{number}" for number in range(14, 39)],
        )
        self.assertEqual(contract["sequenceFieldCount"], 26)
        self.assertEqual(contract["pageFieldCount"], 1)

    def test_replace_text_after_seq_field_preserves_field_nodes(self) -> None:
        """替换标题不得删除SEQ域节点或缓存编号，并应保留换行。"""
        paragraph = ET.fromstring(SEQUENCE_PARAGRAPH_XML)
        namespaces = {"w": W_NS}
        before_fields = len(paragraph.findall(".//w:fldChar", namespaces))
        before_instruction = "".join(
            node.text or ""
            for node in paragraph.findall(".//w:instrText", namespaces)
        )

        MODULE.replace_text_after_seq_field(paragraph, "新标题\n第二行")

        self.assertEqual(
            len(paragraph.findall(".//w:fldChar", namespaces)), before_fields
        )
        self.assertEqual(
            "".join(
                node.text or ""
                for node in paragraph.findall(".//w:instrText", namespaces)
            ),
            before_instruction,
        )
        paragraph_text = "".join(paragraph.itertext())
        self.assertIn("2", paragraph_text)
        self.assertIn("新标题", paragraph_text)
        self.assertIn("第二行", paragraph_text)
        self.assertEqual(len(paragraph.findall(".//w:br", namespaces)), 1)

    def test_replace_sequence_title_reuses_existing_title_paragraphs(self) -> None:
        """双行标题应分写到模板已有段落，不得保留或复制旧第二行。"""
        cell = ET.fromstring(
            f"""
            <w:tc xmlns:w="{W_NS}">
              {SEQUENCE_PARAGRAPH_XML}
              <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>旧第二行</w:t></w:r></w:p>
            </w:tc>
            """
        )

        MODULE.replace_sequence_title(cell, "新第一行\n新第二行")

        paragraphs = cell.findall("w:p", {"w": W_NS})
        self.assertIn("新第一行", "".join(paragraphs[0].itertext()))
        self.assertNotIn("新第二行", "".join(paragraphs[0].itertext()))
        self.assertEqual("".join(paragraphs[1].itertext()), "新第二行")
        self.assertNotIn("旧第二行", "".join(cell.itertext()))

    def test_build_daily_word_report_replaces_only_approved_parts(self) -> None:
        """成品应原位替换25张图与目标文字，其他DOCX部件逐字节不变。"""
        with tempfile.TemporaryDirectory() as directory:
            output_dir = Path(directory)
            for sequence, label, _, _ in MODULE.SMALL_CHART_EXPORT_SPECS:
                Image.new(
                    "RGB", (881, 509), (sequence, sequence, sequence)
                ).save(output_dir / f"{sequence:02d}_{label}.png")
            industry_path = output_dir / "各行业转债正股涨跌幅及估值.png"
            Image.new("RGB", (1762, 1029), "white").save(industry_path)
            two_line_titles = {5, 8, 9, 11, 13, 14, 15, 16, 17, 18, 19, 20}
            chart_titles = [
                f"标题{number}\n第二行"
                if number in two_line_titles
                else f"标题{number}"
                for number in range(1, 25)
            ]

            output_path = MODULE.build_daily_word_report(
                date(2026, 8, 31),
                output_dir,
                _index_frame(),
                chart_titles,
                _industry_frame(),
                industry_path,
            )

            self.assertTrue(output_path.is_file())
            source_hashes = _package_hashes(MODULE.DAILY_WORD_TEMPLATE_PATH)
            output_hashes = _package_hashes(output_path)
            self.assertEqual(source_hashes.keys(), output_hashes.keys())
            changed_parts = {
                name
                for name in source_hashes
                if source_hashes[name] != output_hashes[name]
            }
            expected_changed = {
                "word/document.xml",
                "word/header3.xml",
                *(f"word/media/image{number}.png" for number in range(3, 28)),
            }
            self.assertEqual(changed_parts, expected_changed)

            with zipfile.ZipFile(output_path) as package:
                for part_name in ("word/document.xml", "word/header3.xml"):
                    raw_xml = package.read(part_name).decode("utf-8")
                    declared = set(
                        re.findall(r"xmlns:([A-Za-z0-9]+)=\"[^\"]+\"", raw_xml)
                    )
                    ignorable_match = re.search(
                        r"(?:mc|[^\s=]+):Ignorable=\"([^\"]+)\"", raw_xml
                    )
                    self.assertIsNotNone(ignorable_match)
                    assert ignorable_match is not None
                    missing_prefixes = set(ignorable_match.group(1).split()) - declared
                    self.assertEqual(missing_prefixes, set())
                document = ET.fromstring(package.read("word/document.xml"))
                header = ET.fromstring(package.read("word/header3.xml"))
                namespaces = {"w": W_NS, "r": R_NS, "pr": PKG_REL_NS}
                document_text = "".join(document.itertext())
                header_text = "".join(header.itertext())
                self.assertIn("转债市场日度跟踪20260831", document_text)
                self.assertIn("标题1", document_text)
                self.assertIn("标题5第二行", document_text)
                self.assertIn("标题24", document_text)
                self.assertIn("行业轮动情况：传媒、计算机、煤炭领涨", document_text)
                self.assertIn("2026年08月31日", header_text)
                self.assertEqual(
                    sum(
                        "SEQ 图表" in "".join(
                            node.text or ""
                            for node in paragraph.findall(
                                ".//w:instrText", namespaces
                            )
                        )
                        for paragraph in document.findall(".//w:p", namespaces)
                    ),
                    26,
                )
                body = document.find("w:body", namespaces)
                assert body is not None
                tables = [
                    child
                    for child in body
                    if child.tag == f"{{{W_NS}}}tbl"
                ]
                first_title_cells = tables[2].findall("w:tr", namespaces)[
                    0
                ].findall("w:tc", namespaces)
                last_title_cells = tables[13].findall("w:tr", namespaces)[
                    0
                ].findall("w:tc", namespaces)
                self.assertIn("标题1", "".join(first_title_cells[0].itertext()))
                self.assertIn("标题2", "".join(first_title_cells[2].itertext()))
                self.assertIn("标题24", "".join(last_title_cells[2].itertext()))
                for sequence in range(1, 25):
                    with Image.open(
                        io.BytesIO(
                            package.read(f"word/media/image{sequence + 2}.png")
                        )
                    ) as chart:
                        self.assertEqual(
                            chart.convert("RGB").getpixel((0, 0)),
                            (sequence, sequence, sequence),
                        )
                with Image.open(
                    io.BytesIO(package.read("word/media/image27.png"))
                ) as industry_chart:
                    self.assertEqual(
                        industry_chart.convert("RGB").getpixel((0, 0)),
                        (255, 255, 255),
                    )


if __name__ == "__main__":
    unittest.main()
