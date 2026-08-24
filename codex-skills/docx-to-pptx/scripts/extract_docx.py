"""
extract_docx.py — 从 docx 提取结构化内容（原始 XML 架构，修复版）

架构：原始 XML 解析（zipfile + re + ElementTree）
- 段落提取：正则 <w:p[ >].*?</w:p>（绕过 python-docx 的 mc:AlternateContent 限制）
- Run 格式：正则 <w:r\b[^>]*>（修复：支持带属性的 <w:r w:rsidRPr="...">）
- 图片提取：原始 XML 正则（处理 mc:AlternateContent 嵌套）
- 样式识别：正则 <w:pStyle w:val="..."/>

输出：content.json + content.md + media/
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def extract_runs_from_para(para_xml: str) -> list[dict]:
    """
    从段落 XML 中按 <w:r> 提取每个文本片段及其格式。

    正则 <w:r\b[^>]*> 匹配所有 <w:r> 变体：
    - <w:r>               → 纯标签（旧格式）
    - <w:r w:rsidR="..."> → 带修订属性的标签（Word 修订模式下产生）
    - <w:r w:rsidRPr="..."> → 带格式修订属性

    返回: [{"text": str, "bold": bool, "italic": bool}, ...]
    """
    runs = []
    run_pattern = re.compile(r"<w:r\b[^>]*>(.*?)</w:r>", re.DOTALL)
    for run_xml in run_pattern.findall(para_xml):
        # 加粗: <w:b/> 或 <w:b w:val="1"/>，排除 <w:b w:val="0"/>
        has_b = bool(re.search(r"<w:b[\s/>]", run_xml))
        bold_off = bool(re.search(r'<w:b[^>]*w:val="(0|false)"', run_xml))
        bold = has_b and not bold_off
        # 斜体
        has_i = bool(re.search(r"<w:i[\s/>]", run_xml))
        italic_off = bool(re.search(r'<w:i[^>]*w:val="(0|false)"', run_xml))
        italic = has_i and not italic_off
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", run_xml))
        if text:
            runs.append({"text": text, "bold": bold, "italic": italic})
    return runs


def runs_to_plain(runs: list[dict]) -> str:
    """将 runs 列表拼成纯文本"""
    return "".join(r["text"] for r in runs)


def runs_to_markdown(runs: list[dict]) -> str:
    """将 runs 列表转为 markdown 格式（加粗用 ** 包裹）"""
    result = []
    for r in runs:
        t = r["text"]
        if r["bold"]:
            t = f"**{t}**"
        if r["italic"]:
            t = f"*{t}*"
        result.append(t)
    return "".join(result)


def extract_images_from_docx(docx_path: Path, media_dir: Path) -> dict[str, dict]:
    """从 docx ZIP 提取所有图片，返回 {rId: {filename, path}}"""
    media_dir.mkdir(parents=True, exist_ok=True)
    image_map = {}

    with zipfile.ZipFile(docx_path) as zf:
        rels_xml = zf.read("word/_rels/document.xml.rels")
        rels_root = ET.fromstring(rels_xml)
        for r in rels_root:
            rid = r.get("Id", "")
            target = r.get("Target", "")
            if "image" in r.get("Type", ""):
                src = f"word/{target}"
                if src in zf.namelist():
                    fname = Path(target).name
                    dst = media_dir / fname
                    if not dst.exists():
                        with zf.open(src) as sf, open(dst, "wb") as df:
                            df.write(sf.read())
                    image_map[rid] = {"filename": fname, "path": str(dst)}

    return image_map


def extract(docx_path: Path, output_dir: Path) -> list[dict]:
    """
    主提取函数 → 返回 items 列表

    每个 item：{"type": "heading"|"body"|"chart_title"|"image"|"source",
                "text"?, "level"?, "runs"?, "file"?, "rId"?}
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    media_dir = output_dir / "media"

    # ── 1. 提取图片 ──
    image_map = extract_images_from_docx(docx_path, media_dir)

    # ── 2. 解析 docx XML ──
    with zipfile.ZipFile(docx_path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")

    # ── 3. 提取所有段落 ──
    para_pattern = re.compile(r"<w:p[ >].*?</w:p>", re.DOTALL)
    paras_raw = para_pattern.findall(xml)

    # ── 4. 标题样式映射（模板特定，仅用于 L1/L2/L3 识别）──
    # 文本模式（"图表 N"、"资料来源"）优先于样式 ID，不在此表
    heading_style_ids = {"a0": 1, "a1": 2, "a2": 3}

    # ── 5. 提取所有段落信息 ──
    all_paragraphs = []
    for i, para_xml in enumerate(paras_raw):
        runs = extract_runs_from_para(para_xml)
        text = runs_to_plain(runs).strip()

        # 支持 <w:pStyle w:val="a0"/> 和 <w:pStyle w:val="a0" /> 两种格式
        style_match = re.search(r'<w:pStyle w:val="([^"]+)"', para_xml)
        style = style_match.group(1) if style_match else ""

        blips = re.findall(r'<a:blip[^>]*r:embed="(rId\d+)"', para_xml)
        images = [{"rId": rid, "file": image_map.get(rid, {}).get("path", "")}
                  for rid in blips if rid in image_map]

        if not text and not images:
            continue

        all_paragraphs.append({
            "index": i,
            "text": text,
            "runs": runs,
            "style": style,
            "images": images,
        })

    # ── 6. 找第一个 L1 标题 ──
    first_l1_idx = None
    for para in all_paragraphs:
        if para["style"] == "a0":
            first_l1_idx = para["index"]
            break

    if first_l1_idx is None:
        print("[ERROR] 未找到一级标题（a0 样式）", file=sys.stderr)
        return []

    # ── 7. 找"风险提示"章 ──
    risk_para_idx = None
    risk_content_idx = None
    for j, para in enumerate(all_paragraphs):
        if para["style"] == "a0" and "风险提示" in para["text"]:
            risk_para_idx = para["index"]
            if j + 1 < len(all_paragraphs):
                risk_content_idx = all_paragraphs[j + 1]["index"]
            break

    # ── 8. 过滤：first L1 ~ risk 之后第一段 ──
    filtered = []
    for para in all_paragraphs:
        if para["index"] < first_l1_idx:
            continue
        if risk_content_idx and para["index"] > risk_content_idx:
            continue
        if para["style"] in ("TOC1", "TOC2", "TOC3"):
            continue
        filtered.append(para)

    # ── 9. 生成 items（文本模式优先，样式 ID 仅用于标题识别）──
    items = []
    skip_until_next_heading = False

    for para in filtered:
        text = para["text"]
        runs = para.get("runs", [])
        style = para["style"]
        images = para["images"]

        # ── 第一优先级：文本模式（通用，不依赖任何模板样式 ID）──
        if re.match(r"^图表\s*\d+", text):
            if skip_until_next_heading:
                continue
            items.append({"type": "chart_title", "text": text})
        elif re.match(r"^资料来源[：:]", text):
            if skip_until_next_heading:
                continue
            items.append({"type": "source", "text": text})
        elif re.match(r"^注[：:]", text):
            continue  # 图表附注，静默跳过

        # ── 第二优先级：样式 ID → 标题级别（模板特定）──
        elif style in heading_style_ids:
            level = heading_style_ids[style]
            if level == 1:
                if "复盘" in text:
                    skip_until_next_heading = True
                    continue
                else:
                    skip_until_next_heading = False
            if skip_until_next_heading:
                continue
            items.append({"type": "heading", "level": level, "text": text, "runs": runs})

        # ── 第三优先级：其他 → 正文 ──
        elif len(text) > 5:
            if skip_until_next_heading:
                continue
            items.append({"type": "body", "text": text, "runs": runs})

        if images and not skip_until_next_heading:
            for img in images:
                items.append({"type": "image", "rId": img["rId"], "file": img["file"]})

    # ── 10. 生成 markdown ──
    md_lines = ["# docx 内容提取结果\n"]
    md_lines.append(f"**总计**: {len(items)} 项\n")
    md_lines.append(f"- 一级标题: {sum(1 for i in items if i['type'] == 'heading' and i.get('level') == 1)} 个")
    md_lines.append(f"- 二级标题: {sum(1 for i in items if i['type'] == 'heading' and i.get('level') == 2)} 个")
    md_lines.append(f"- 三级标题: {sum(1 for i in items if i['type'] == 'heading' and i.get('level') == 3)} 个")
    md_lines.append(f"- 正文段落: {sum(1 for i in items if i['type'] == 'body')} 个")
    md_lines.append(f"- 图表标题: {sum(1 for i in items if i['type'] == 'chart_title')} 个")
    md_lines.append(f"- 图片: {sum(1 for i in items if i['type'] == 'image')} 个\n")
    md_lines.append("---\n")

    for item in items:
        t = item["type"]
        text = item.get("text", "")
        runs = item.get("runs", [])
        if t == "heading":
            level = item.get("level", 2)
            md_text = runs_to_markdown(runs) if runs else text
            md_lines.append(f"{'#' * level} {md_text}\n")
        elif t == "chart_title":
            md_lines.append(f"> **📊 {text}**\n")
        elif t == "body":
            md_text = runs_to_markdown(runs) if runs else text
            md_lines.append(f"{md_text}\n")
        elif t == "image":
            src = item.get("file", "")
            md_lines.append(f"![图片]({src})\n")
        elif t == "source":
            md_lines.append(f"*{text}*\n")

    with open(output_dir / "content.md", "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    with open(output_dir / "content.json", "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)

    # ── 11. 统计 ──
    n_headings = sum(1 for i in items if i["type"] == "heading")
    n_body = sum(1 for i in items if i["type"] == "body")
    n_chart = sum(1 for i in items if i["type"] == "chart_title")
    n_images = sum(1 for i in items if i["type"] == "image")

    print(f"[OK] {output_dir}")
    print(f"     标题: {n_headings}, 段落: {n_body}, 图表: {n_chart}, 图片: {n_images}")
    print(f"     markdown: {output_dir / 'content.md'}")

    return items


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    docx_path = Path(args.input)
    if not docx_path.exists():
        print(f"[ERROR] {docx_path}", file=sys.stderr)
        return 1

    items = extract(docx_path, Path(args.output))
    if not items:
        print("[ERROR] 未提取到可用内容；请检查 a0/a1/a2 标题样式与文档结构", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
