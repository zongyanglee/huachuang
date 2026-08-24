from __future__ import annotations

import copy
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


def first_drawing_run(document_xml: bytes) -> etree._Element:
    root = etree.fromstring(document_xml)
    drawing = root.xpath('//*[local-name()="drawing"]')[0]
    return drawing.getparent()


def main() -> None:
    workspace = Path(__file__).resolve().parents[2].parent
    original = next(p for p in workspace.glob("*20260810.docx") if "_" not in p.name)
    updated = next(p for p in workspace.glob("*20260810_*.docx"))

    with zipfile.ZipFile(original) as source:
        original_document = source.read("word/document.xml")
        original_chart = source.read("word/charts/chart1.xml")
        original_chart_rels = source.read("word/charts/_rels/chart1.xml.rels")
        chart_run = first_drawing_run(original_document)

    with zipfile.ZipFile(updated) as current:
        document_root = etree.fromstring(current.read("word/document.xml"))
        current_drawing = document_root.xpath('//*[local-name()="drawing"]')[0]
        image_run = current_drawing.getparent()
        image_run.getparent().replace(image_run, copy.deepcopy(chart_run))
        restored_document = etree.tostring(
            document_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        rels_root = etree.fromstring(current.read("word/_rels/document.xml.rels"))
        removed_targets = []
        for relationship in list(rels_root):
            if relationship.get("Id") == "rId23":
                removed_targets.append("word/" + relationship.get("Target"))
                rels_root.remove(relationship)
        restored_rels = etree.tostring(
            rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        replacements = {
            "word/document.xml": restored_document,
            "word/_rels/document.xml.rels": restored_rels,
            "word/charts/chart1.xml": original_chart,
            "word/charts/_rels/chart1.xml.rels": original_chart_rels,
        }
        with tempfile.NamedTemporaryFile(
            dir=updated.parent, prefix="restore_figure1_", suffix=".docx", delete=False
        ) as handle:
            temp_path = Path(handle.name)
        try:
            with zipfile.ZipFile(temp_path, "w") as target:
                for item in current.infolist():
                    if item.filename in removed_targets:
                        continue
                    target.writestr(item, replacements.get(item.filename, current.read(item.filename)))
            with zipfile.ZipFile(temp_path) as audit:
                if audit.testzip() is not None:
                    raise RuntimeError("恢复后的 DOCX 压缩包校验失败")
            shutil.move(str(temp_path), str(updated))
        finally:
            if temp_path.exists():
                temp_path.unlink()
    print(f"已将 {updated.name} 的图1恢复为原模板图表")


if __name__ == "__main__":
    main()
