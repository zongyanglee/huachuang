from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from lxml import etree


C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"c": C_NS, "a": A_NS}
Q = lambda namespace, tag: f"{{{namespace}}}{tag}"

SERIES = {
    "中证转债": {"parquet_name": "转债指数", "column": "B", "color": "E6121B"},
    "万得全A": {"parquet_name": "万得全A", "column": "C", "color": "0262BA"},
}


def excel_serial(ts: pd.Timestamp) -> int:
    return (ts.normalize() - pd.Timestamp("1899-12-30")).days


def find_parquet_root(workspace: Path) -> Path:
    candidates = [
        p
        for p in workspace.iterdir()
        if p.is_dir() and p.name.endswith("历史序列") and list(p.glob("202[3-9]/**/*.parquet"))
    ]
    if len(candidates) != 1:
        raise RuntimeError(f"无法唯一确定 parquet 目录，候选为：{candidates}")
    return candidates[0]


def load_index_data(parquet_root: Path) -> pd.DataFrame:
    target_names = {item["parquet_name"] for item in SERIES.values()}
    records: list[tuple[pd.Timestamp, str, float]] = []

    parquet_files = sorted(parquet_root.glob("202[3-9]/**/*.parquet"))
    if not parquet_files:
        raise RuntimeError(f"未在 {parquet_root} 找到2023年以来的 parquet 文件")

    for parquet_path in parquet_files:
        frame = pd.read_parquet(parquet_path)
        subset = frame.loc[
            (frame["__sheet_name"] == "指数") & frame["__row_id"].isin(target_names)
        ]
        if subset.empty:
            continue
        for _, row in subset.iterrows():
            series_name = str(row["__row_id"])
            for column in frame.columns[2:]:
                date = pd.to_datetime(column, errors="coerce")
                if pd.isna(date) or date < pd.Timestamp("2023-01-01"):
                    continue
                value = pd.to_numeric(row[column], errors="coerce")
                if pd.notna(value):
                    records.append((date.normalize(), series_name, float(value)))

    long = pd.DataFrame(records, columns=["date", "series", "value"])
    if long.empty:
        raise RuntimeError("没有读到图1所需的指数数据")
    long = long.drop_duplicates(["date", "series"], keep="last")
    wide = long.pivot(index="date", columns="series", values="value").sort_index()
    wide = wide[["转债指数", "万得全A"]].dropna()
    if wide.empty:
        raise RuntimeError("中证转债与万得全A没有可共同绘制的日期")
    return wide


def set_num_cache(cache: etree._Element, values: list[str]) -> None:
    point_count = cache.find("c:ptCount", NS)
    if point_count is None:
        point_count = etree.SubElement(cache, Q(C_NS, "ptCount"))
    point_count.set("val", str(len(values)))
    for point in cache.findall("c:pt", NS):
        cache.remove(point)
    for idx, value in enumerate(values):
        point = etree.SubElement(cache, Q(C_NS, "pt"), idx=str(idx))
        etree.SubElement(point, Q(C_NS, "v")).text = value


def set_series_color(series: etree._Element, color: str) -> None:
    color_node = series.find("c:spPr/a:ln/a:solidFill/a:srgbClr", NS)
    if color_node is None:
        raise RuntimeError("图表序列缺少预期的折线颜色节点")
    color_node.set("val", color)


def update_chart_xml(chart_xml: bytes, data: pd.DataFrame) -> bytes:
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(chart_xml, parser)
    date_values = [str(excel_serial(ts)) for ts in data.index]

    found: set[str] = set()
    for series in root.findall(".//c:lineChart/c:ser", NS):
        name = "".join(v.text or "" for v in series.findall(".//c:tx//c:v", NS)).strip()
        if name not in SERIES:
            continue
        spec = SERIES[name]
        parquet_name = spec["parquet_name"]
        value_values = [format(value, ".15g") for value in data[parquet_name].tolist()]

        category_cache = series.find("c:cat/c:numRef/c:numCache", NS)
        value_cache = series.find("c:val/c:numRef/c:numCache", NS)
        if category_cache is None or value_cache is None:
            raise RuntimeError(f"序列{name}缺少数值缓存")
        set_num_cache(category_cache, date_values)
        set_num_cache(value_cache, value_values)
        set_series_color(series, spec["color"])

        formulas = series.findall(".//c:f", NS)
        if len(formulas) >= 3:
            start_row = 1462
            end_row = start_row + len(data) - 1
            formulas[1].text = f"'可转债-指数行情'!$A${start_row}:$A${end_row}"
            column = spec["column"]
            formulas[2].text = (
                f"'可转债-指数行情'!${column}${start_row}:${column}${end_row}"
            )
        found.add(name)

    missing = set(SERIES) - found
    if missing:
        raise RuntimeError(f"图1中未找到序列：{sorted(missing)}")

    # 图表数据已完整写入 OOXML 缓存。移除已失效的旧 Excel 外链，避免 Word
    # 每次打开或导出时尝试访问原作者电脑上的 E 盘文件。
    for external_data in root.findall("c:externalData", NS):
        root.remove(external_data)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def remove_external_chart_relationships(rels_xml: bytes) -> bytes:
    root = etree.fromstring(rels_xml)
    for relationship in list(root):
        rel_type = relationship.get("Type", "")
        if rel_type.endswith("/oleObject") and relationship.get("TargetMode") == "External":
            root.remove(relationship)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_docx(docx_path: Path, replacements: dict[str, bytes]) -> None:
    with tempfile.NamedTemporaryFile(
        dir=docx_path.parent, prefix="figure1_", suffix=".docx", delete=False
    ) as handle:
        temp_path = Path(handle.name)
    try:
        with zipfile.ZipFile(docx_path, "r") as source, zipfile.ZipFile(
            temp_path, "w"
        ) as target:
            for item in source.infolist():
                payload = replacements.get(item.filename, source.read(item.filename))
                target.writestr(item, payload)
        with zipfile.ZipFile(temp_path, "r") as audit:
            bad = audit.testzip()
            if bad:
                raise RuntimeError(f"更新后的 DOCX 压缩包损坏：{bad}")
        shutil.move(str(temp_path), str(docx_path))
    finally:
        if temp_path.exists():
            temp_path.unlink()


def audit_chart(docx_path: Path) -> None:
    with zipfile.ZipFile(docx_path) as archive:
        root = etree.fromstring(archive.read("word/charts/chart1.xml"))
    for series in root.findall(".//c:lineChart/c:ser", NS):
        name = "".join(v.text or "" for v in series.findall(".//c:tx//c:v", NS)).strip()
        if name not in SERIES:
            continue
        cats = series.findall("c:cat/c:numRef/c:numCache/c:pt", NS)
        vals = series.findall("c:val/c:numRef/c:numCache/c:pt", NS)
        first_date = datetime(1899, 12, 30) + pd.Timedelta(days=int(cats[0].findtext("c:v", namespaces=NS)))
        last_date = datetime(1899, 12, 30) + pd.Timedelta(days=int(cats[-1].findtext("c:v", namespaces=NS)))
        color = series.find("c:spPr/a:ln/a:solidFill/a:srgbClr", NS).get("val")
        print(
            f"{name}: {len(vals)}点, {first_date:%Y-%m-%d}至{last_date:%Y-%m-%d}, "
            f"首值{vals[0].findtext('c:v', namespaces=NS)}, "
            f"末值{vals[-1].findtext('c:v', namespaces=NS)}, 颜色#{color}"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="更新人保转债周报图1的两条指数时序线")
    parser.add_argument("docx", type=Path)
    args = parser.parse_args()
    docx_path = args.docx.resolve()
    if not docx_path.exists():
        raise FileNotFoundError(docx_path)

    parquet_root = find_parquet_root(docx_path.parent)
    data = load_index_data(parquet_root)
    with zipfile.ZipFile(docx_path, "r") as archive:
        chart_xml = archive.read("word/charts/chart1.xml")
        rels_name = "word/charts/_rels/chart1.xml.rels"
        rels_xml = archive.read(rels_name)
    patch_docx(
        docx_path,
        {
            "word/charts/chart1.xml": update_chart_xml(chart_xml, data),
            rels_name: remove_external_chart_relationships(rels_xml),
        },
    )

    print(
        f"数据区间：{data.index.min():%Y-%m-%d}至{data.index.max():%Y-%m-%d}，"
        f"共同交易日{len(data)}个"
    )
    audit_chart(docx_path)


if __name__ == "__main__":
    main()
