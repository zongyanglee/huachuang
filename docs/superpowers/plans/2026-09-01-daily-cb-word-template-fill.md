# Daily Convertible-Bond Word Template Fill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend the existing convertible-bond daily script so every successful run fills the frozen Word template with the current date, index table, 24 numbered charts, 24 chart titles, and the industry table/title.

**Architecture:** Keep the Word file as the design authority. Copy its ZIP package, validate the frozen structure, patch only approved WordprocessingML text nodes and the 25 existing media parts, then atomically publish a new `.docx`; never rebuild the document or import another project script.

**Tech Stack:** Python 3.12 standard library (`zipfile`, `xml.etree.ElementTree`, `hashlib`, `shutil`), pandas, Pillow, unittest, Microsoft Word only for development-time PDF rendering.

**Spec:** `docs/superpowers/specs/2026-09-01-daily-cb-word-template-fill-design.md`

## Global Constraints

- Modify `src/daily/【日报】转债日报.py`; do not create an imported project module.
- Keep `D:\JupyterFiles\huachuang\【华创固收】转债市场日度跟踪20260831.docx` byte-for-byte unchanged.
- Require template SHA-256 `922AA1FC6DA6C384E264A7077CBB7CE02FAE9C7501539D4E8B74CC76C4303FB3`.
- Preserve 1 section, 18 top-level body tables, 25 daily chart image relationships, 26 `SEQ 图表` fields, 1 `PAGE` field, all content controls, and every floating anchor.
- Do not update the report summary or any compliance, analyst, team, contact, rating, source, or disclaimer content.
- Output `runs/daily/YYYYMMDD_转债日报/【华创固收】转债市场日度跟踪YYYYMMDD.docx` atomically.
- Runtime Word COM is prohibited; Word is used only to render the validation sample.

---

### Task 1: Add deterministic Word display-data builders

**Files:**
- Modify: `src/daily/【日报】转债日报.py:40-2300`
- Create: `tests/test_cb_daily_word_template.py`

**Interfaces:**
- Consumes: `MAIN_INDEX_SPECS`, `STYLE_INDEX_SPECS`, `pd.DataFrame index_performance`, `pd.DataFrame industry_performance`.
- Produces: `build_word_index_table_rows(index_performance: pd.DataFrame) -> list[list[str]]` and `build_industry_rotation_title(industry_performance: pd.DataFrame) -> str`.

- [ ] **Step 1: Write failing tests for index table rows and industry title**

```python
def test_build_word_index_table_rows_formats_both_groups() -> None:
    records = []
    for group, specs in (
        ("主要指数", MODULE.MAIN_INDEX_SPECS),
        ("风格指数", MODULE.STYLE_INDEX_SPECS),
    ):
        for offset, (_, _, display_name) in enumerate(specs):
            records.append({
                "组别": group,
                "指数名称": display_name,
                "收盘价": 100 + offset,
                "日涨跌幅": -0.125 + offset,
                "近一周涨跌幅": 1 + offset,
                "近一月涨跌幅": 2 + offset,
                "年初至今涨跌幅": 3 + offset,
            })
    rows = MODULE.build_word_index_table_rows(pd.DataFrame(records))
    assert len(rows) == 9
    assert rows[0] == [
        "中证转债", "100.00", "-0.12", "1.00", "2.00", "3.00",
        "大盘指数", "100.00", "-0.12", "1.00", "2.00", "3.00",
    ]


def test_build_industry_rotation_title_uses_top_three_stock_returns() -> None:
    data = pd.DataFrame({
        "行业名称": ["煤炭", "传媒", "计算机", "电子"],
        "正股日涨跌幅": [2.1, 3.8, 2.2, -1.0],
    })
    assert MODULE.build_industry_rotation_title(data) == (
        "行业轮动情况：传媒、计算机、煤炭领涨"
    )
```

- [ ] **Step 2: Run the focused tests and verify missing-function failures**

Run: `python -m unittest tests.test_cb_daily_word_template.DailyWordTemplateTests.test_build_word_index_table_rows_formats_both_groups tests.test_cb_daily_word_template.DailyWordTemplateTests.test_build_industry_rotation_title_uses_top_three_stock_returns -v`

Expected: FAIL because both functions are absent.

- [ ] **Step 3: Implement strict builders**

```python
WORD_INDEX_VALUE_COLUMNS = (
    "收盘价", "日涨跌幅", "近一周涨跌幅", "近一月涨跌幅", "年初至今涨跌幅",
)
WORD_INDEX_DISPLAY_NAMES = {
    "可转债等权": "转债等权",
    "可转债正股等权": "正股等权",
    "可转债预案": "转债预案",
    "大盘指数(申万)": "大盘指数",
    "中盘指数(申万)": "中盘指数",
    "小盘指数(申万)": "小盘指数",
}


def build_word_index_table_rows(index_performance: pd.DataFrame) -> list[list[str]]:
    required = {"组别", "指数名称", *WORD_INDEX_VALUE_COLUMNS}
    missing = required - set(index_performance.columns)
    if missing:
        raise RuntimeError(f"Word指数表缺少字段：{sorted(missing)}")
    rows_by_name = index_performance.set_index("指数名称", drop=False)
    main_names = [display for _, _, display in MAIN_INDEX_SPECS]
    style_names = [display for _, _, display in STYLE_INDEX_SPECS]
    missing_names = [name for name in (*main_names, *style_names) if name not in rows_by_name.index]
    if missing_names:
        raise RuntimeError(f"Word指数表缺少指数：{missing_names}")
    result = []
    for main_name, style_name in zip(main_names, style_names):
        row = []
        for name in (main_name, style_name):
            values = rows_by_name.loc[name]
            word_name = WORD_INDEX_DISPLAY_NAMES.get(name, name)
            row.extend([word_name, *(f"{float(values[column]):.2f}" for column in WORD_INDEX_VALUE_COLUMNS)])
        result.append(row)
    return result


def build_industry_rotation_title(industry_performance: pd.DataFrame) -> str:
    required = {"行业名称", "正股日涨跌幅"}
    missing = required - set(industry_performance.columns)
    if missing:
        raise RuntimeError(f"行业轮动标题缺少字段：{sorted(missing)}")
    ranked = industry_performance.dropna(subset=list(required)).copy()
    ranked["_order"] = range(len(ranked))
    ranked = ranked.sort_values(["正股日涨跌幅", "_order"], ascending=[False, True])
    if len(ranked) < 3:
        raise RuntimeError("行业轮动标题至少需要3个有效行业")
    return "行业轮动情况：" + "、".join(ranked.head(3)["行业名称"]) + "领涨"
```

- [ ] **Step 4: Run the focused tests and verify PASS**

Run: `python -m unittest tests.test_cb_daily_word_template -v`

Expected: both builder tests PASS.

- [ ] **Step 5: Commit the builders**

```bash
git add src/daily/【日报】转债日报.py tests/test_cb_daily_word_template.py
git commit -m "feat: build daily Word report display data"
```

---

### Task 2: Add frozen-template validation and XML text patching

**Files:**
- Modify: `src/daily/【日报】转债日报.py:1-80`
- Modify: `src/daily/【日报】转债日报.py:7260-7340`
- Modify: `tests/test_cb_daily_word_template.py`

**Interfaces:**
- Consumes: `Path template_path`, parsed `word/document.xml`, `word/header3.xml`, and `word/_rels/document.xml.rels`.
- Produces: `inspect_daily_word_template(template_path: Path) -> dict[str, object]`, `set_plain_text_control_value(root: ET.Element, tag: str, value: str) -> int`, and `replace_text_after_seq_field(paragraph: ET.Element, value: str) -> None`.

- [ ] **Step 1: Write failing structure and field-preservation tests**

```python
def test_template_contract_matches_frozen_reference() -> None:
    contract = MODULE.inspect_daily_word_template(MODULE.DAILY_WORD_TEMPLATE_PATH)
    assert contract["sha256"] == MODULE.DAILY_WORD_TEMPLATE_SHA256
    assert contract["topLevelTableCount"] == 18
    assert contract["chartImageRelationshipIds"] == [f"rId{i}" for i in range(14, 39)]
    assert contract["sequenceFieldCount"] == 26
    assert contract["pageFieldCount"] == 1


def test_replace_text_after_seq_field_preserves_field_nodes() -> None:
    paragraph = ET.fromstring(SEQUENCE_PARAGRAPH_XML)
    before = len(paragraph.findall(".//w:fldChar", MODULE.WORD_XML_NAMESPACES))
    MODULE.replace_text_after_seq_field(paragraph, "新标题\n第二行")
    after = len(paragraph.findall(".//w:fldChar", MODULE.WORD_XML_NAMESPACES))
    assert before == after == 3
    assert "新标题" in "".join(paragraph.itertext())
    assert len(paragraph.findall(".//w:br", MODULE.WORD_XML_NAMESPACES)) == 1
```

- [ ] **Step 2: Run tests and verify failure at absent APIs**

Run: `python -m unittest tests.test_cb_daily_word_template -v`

Expected: FAIL on `inspect_daily_word_template` and XML helper lookups.

- [ ] **Step 3: Implement namespaces, constants, structure inspection, and targeted text helpers**

```python
import hashlib
import shutil

DAILY_WORD_TEMPLATE_PATH = WORKSPACE / "【华创固收】转债市场日度跟踪20260831.docx"
DAILY_WORD_TEMPLATE_SHA256 = "922AA1FC6DA6C384E264A7077CBB7CE02FAE9C7501539D4E8B74CC76C4303FB3"
WORD_XML_NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def inspect_daily_word_template(template_path: Path) -> dict[str, object]:
    digest = hashlib.sha256(template_path.read_bytes()).hexdigest().upper()
    with zipfile.ZipFile(template_path) as package:
        document_root = ET.fromstring(package.read("word/document.xml"))
        body = document_root.find("w:body", WORD_XML_NAMESPACES)
        tables = [child for child in body if child.tag == _word_tag("tbl")]
        relationship_ids = []
        for table in tables[2:15]:
            relationship_ids.extend(
                blip.get(_rel_tag("embed"))
                for blip in table.findall(".//a:blip", WORD_XML_NAMESPACES)
            )
        return {
            "sha256": digest,
            "topLevelTableCount": len(tables),
            "chartImageRelationshipIds": relationship_ids,
            "sequenceFieldCount": sum(
                "SEQ 图表" in (node.text or "")
                for node in document_root.findall(".//w:instrText", WORD_XML_NAMESPACES)
            ),
            "pageFieldCount": _count_fields(package, "PAGE"),
        }
```

`replace_text_after_seq_field()` must find the `w:fldCharType="end"` run, preserve that run and all prior field runs, reuse the first subsequent run's `w:rPr`, remove only subsequent title-content runs, then append one text node per line separated by `w:br`. `set_plain_text_control_value()` must select `w:sdt` by `w:sdtPr/w:tag/@w:val`, update only its first `w:t`, and blank remaining cached text nodes.

- [ ] **Step 4: Run tests and verify PASS**

Run: `python -m unittest tests.test_cb_daily_word_template -v`

Expected: all Task 1 and Task 2 tests PASS.

- [ ] **Step 5: Commit template validation and XML helpers**

```bash
git add src/daily/【日报】转债日报.py tests/test_cb_daily_word_template.py
git commit -m "feat: validate and patch daily Word template XML"
```

---

### Task 3: Build the atomic DOCX package writer

**Files:**
- Modify: `src/daily/【日报】转债日报.py:7260-7340`
- Modify: `tests/test_cb_daily_word_template.py`

**Interfaces:**
- Consumes: template path, run date, index rows, 24 titles, industry title, 24 numbered chart paths, and one industry image path.
- Produces: `build_daily_word_report(run_date: date, output_dir: Path, index_performance: pd.DataFrame, chart_titles: list[str], industry_performance: pd.DataFrame, industry_chart_path: Path, template_path: Path = DAILY_WORD_TEMPLATE_PATH) -> Path`.

- [ ] **Step 1: Write failing end-to-end package test**

```python
def test_build_daily_word_report_replaces_approved_slots_only() -> None:
    with tempfile.TemporaryDirectory() as directory:
        output_dir = Path(directory)
        for sequence, label, _, _ in MODULE.SMALL_CHART_EXPORT_SPECS:
            Image.new("RGB", (881, 509), (sequence, 0, 0)).save(
                output_dir / f"{sequence:02d}_{label}.png"
            )
        industry_path = output_dir / "各行业转债正股涨跌幅及估值.png"
        Image.new("RGB", (1762, 1029), "white").save(industry_path)
        output_path = MODULE.build_daily_word_report(
            date(2026, 8, 31), output_dir, INDEX_FRAME,
            [f"标题{i}" for i in range(1, 25)], INDUSTRY_FRAME,
            industry_path,
        )
        assert output_path.is_file()
        assert MODULE._sha256(MODULE.DAILY_WORD_TEMPLATE_PATH) == MODULE.DAILY_WORD_TEMPLATE_SHA256
        with zipfile.ZipFile(output_path) as package:
            root = ET.fromstring(package.read("word/document.xml"))
            assert len(root.findall(".//w:instrText", MODULE.WORD_XML_NAMESPACES)) >= 26
            assert package.read("word/media/image3.png") != package.read("word/media/image4.png")
```

- [ ] **Step 2: Run the package test and verify missing writer failure**

Run: `python -m unittest tests.test_cb_daily_word_template.DailyWordTemplateTests.test_build_daily_word_report_replaces_approved_slots_only -v`

Expected: FAIL because `build_daily_word_report` is absent.

- [ ] **Step 3: Implement the atomic package writer**

```python
def build_daily_word_report(
    run_date: date,
    output_dir: Path,
    index_performance: pd.DataFrame,
    chart_titles: list[str],
    industry_performance: pd.DataFrame,
    industry_chart_path: Path,
    template_path: Path = DAILY_WORD_TEMPLATE_PATH,
) -> Path:
    if len(chart_titles) != 24:
        raise RuntimeError(f"Word图表标题数量异常：{len(chart_titles)}")
    contract = validate_daily_word_template(template_path)
    chart_paths = [output_dir / f"{n:02d}_{label}.png" for n, label, _, _ in SMALL_CHART_EXPORT_SPECS]
    missing = [str(path) for path in (*chart_paths, industry_chart_path) if not path.is_file()]
    if missing:
        raise FileNotFoundError(f"Word报告缺少图片：{missing}")
    output_path = output_dir / f"【华创固收】转债市场日度跟踪{run_date:%Y%m%d}.docx"
    temporary_path = output_path.with_name(f".{output_path.stem}.{os.getpid()}.tmp.docx")
    try:
        _write_patched_word_package(
            template_path, temporary_path, run_date,
            build_word_index_table_rows(index_performance), chart_titles,
            build_industry_rotation_title(industry_performance),
            [*chart_paths, industry_chart_path], contract,
        )
        validate_generated_daily_word_report(temporary_path, run_date)
        os.replace(temporary_path, output_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    return output_path
```

`_write_patched_word_package()` must copy every original ZIP member with its original `ZipInfo`; only `word/document.xml`, `word/header3.xml`, and the 25 relationship-resolved media members receive replacement bytes. It must register source namespace prefixes before serialization. `validate_generated_daily_word_report()` must verify ZIP integrity, output date text, 18 top-level tables, 25 relationships, 26 sequence fields, all target titles, all 9 index rows, and unchanged part-name inventory.

- [ ] **Step 4: Run Word-specific tests and verify PASS**

Run: `python -m unittest tests.test_cb_daily_word_template -v`

Expected: all Word-template tests PASS.

- [ ] **Step 5: Commit the package writer**

```bash
git add src/daily/【日报】转债日报.py tests/test_cb_daily_word_template.py
git commit -m "feat: fill frozen daily Word report template"
```

---

### Task 4: Integrate Word generation into the daily run

**Files:**
- Modify: `src/daily/【日报】转债日报.py:8635-9250`
- Modify: `tests/test_cb_daily_word_template.py`

**Interfaces:**
- Consumes: the existing `long_chart_titles`, `index_performance`, `industry_performance`, `industry_performance_png`, `run_date`, and `output_dir` values in `run()`.
- Produces: one Word report per run and metadata key `Word报告`.

- [ ] **Step 1: Write failing integration-order and metadata tests**

```python
def test_run_source_calls_word_builder_after_numbered_charts() -> None:
    source = SCRIPT.read_text(encoding="utf-8")
    small_pos = source.index("export_numbered_titleless_small_charts(output_dir)")
    word_pos = source.index("build_daily_word_report(", small_pos)
    assert word_pos > small_pos
    assert '"Word报告": str(word_report_path)' in source
```

- [ ] **Step 2: Run the integration test and verify failure**

Run: `python -m unittest tests.test_cb_daily_word_template.DailyWordTemplateTests.test_run_source_calls_word_builder_after_numbered_charts -v`

Expected: FAIL because the call and metadata key are absent.

- [ ] **Step 3: Add run-stage progress, output path, builder call, and metadata**

```python
report_progress(97, "填充 Word 报告")
word_report_path = build_daily_word_report(
    run_date,
    output_dir,
    index_performance,
    long_chart_titles,
    industry_performance,
    industry_performance_png,
)
```

Add `"Word报告": str(word_report_path)` to the returned metadata. Keep Word generation after the 24 titleless charts and `long_chart_titles` exist, and before final cleanup/return.

- [ ] **Step 4: Run integration and existing focused suites**

Run: `python -m unittest tests.test_cb_daily_word_template tests.test_cb_daily_small_chart_exports tests.test_cb_daily_commentary tests.test_cb_daily_overview_layout -v`

Expected: PASS with no existing output-order regression.

- [ ] **Step 5: Commit run integration**

```bash
git add src/daily/【日报】转债日报.py tests/test_cb_daily_word_template.py
git commit -m "feat: emit Word report from daily workflow"
```

---

### Task 5: Generate and visually verify the 20260831 sample

**Files:**
- Output: `runs/daily/20260831_转债日报/【华创固收】转债市场日度跟踪20260831.docx`
- QA only: `tmp/word_template_audit_20260831/final-render/`

**Interfaces:**
- Consumes: existing 20260831 numbered charts, industry table image, template, and data required by the Word-only validation entry point.
- Produces: validated DOCX and 12-page Word-native render evidence.

- [ ] **Step 1: Run syntax and full automated tests**

Run: `python tests/validate_python_syntax.py`

Expected: PASS.

Run: `python -m unittest discover -s tests -p "test_cb_daily_*.py" -v`

Expected: PASS.

- [ ] **Step 2: Generate the 20260831 Word sample from existing outputs**

Create and run the QA-only fixture `tmp/word_template_audit_20260831/build_word_sample.py`. It imports the daily module, reconstructs `index_performance` and `industry_performance` from `转债日报市场数据底稿_20260831.xlsx`, parses the 24 titles after `图表标题：` in `转债日报点评_20260831.txt`, and calls `build_daily_word_report()` with the existing 24 numbered PNGs and industry PNG. It must not call iFinD or Wind and must not add a runtime CLI mode.

Expected: `【华创固收】转债市场日度跟踪20260831.docx` exists in the 20260831 output directory and the source template hash is unchanged.

- [ ] **Step 3: Export the generated DOCX through Word in STA mode**

Run: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File tmp/word_template_audit_20260831/audit_word.ps1 -InputDocx <generated-docx> -OutputPdf <generated-pdf>`

Expected: `Pages=12; Tables=18; InlineShapes=25` and a non-empty PDF.

- [ ] **Step 4: Render and inspect every PDF page**

Run: `pdftoppm -png -r 144 <generated-pdf> tmp/word_template_audit_20260831/final-render/page`

Expected: 12 PNG pages. Inspect all pages at 100%: page 1 date only, pages 2—7 updated table/charts/titles without clipping or pagination shifts, pages 8—12 unchanged.

- [ ] **Step 5: Compare structural inventories and hashes**

Run the Word test helper to compare source and output ZIP member names, preserve-only member SHA-256 values, table/field/control counts, and relationship targets.

Expected: only `word/document.xml`, `word/header3.xml`, and the 25 approved media parts differ.

- [ ] **Step 6: Commit verification fixtures or final corrections**

```bash
git add src/daily/【日报】转债日报.py tests/test_cb_daily_word_template.py
git commit -m "test: verify daily Word template fidelity"
```
