"""
plan_slides.py — 基于 markdown 语义分段规划 PPT 页面（V3）

核心逻辑：
    1. 跳过目录部分（只有标题没有正文的区域）
    2. 按 L1 标题分大节，每节内按图片数量分页
    3. 正文过多时自动拆分
    4. 图表标题在图片上方
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


def parse_content(items: list[dict]) -> list[dict]:
    """从 content.json 的 items 列表解析出 sections
    
    跳过目录部分（只有标题没有正文的区域）
    每个 section = 一个 L1 标题下的所有内容
    """
    # 1. 找到实际内容的起点：第一个 body 类型且在第一个 image 之前
    first_image_idx = None
    for i, item in enumerate(items):
        if item["type"] == "image":
            first_image_idx = i
            break
    
    if first_image_idx is None:
        return []
    
    # 2. 从 first_image_idx 向前找，找第一个 body 类型（实际内容的起点）
    start_idx = 0
    for i in range(first_image_idx - 1, -1, -1):
        if items[i]["type"] == "body":
            start_idx = i
            break
    
    # 3. 从 start_idx 开始解析
    sections: list[dict] = []
    current: dict | None = None
    
    for item in items[start_idx:]:
        t = item["type"]
        
        if t == "heading":
            level = item.get("level", 2)
            text = item["text"]
            
            # L1 标题：开始新 section
            if level == 1:
                if current and (current["body"] or current["charts"]):
                    sections.append(current)
                current = {
                    "heading": text,
                    "body": [],
                    "charts": [],
                }
            # L2/L3 标题：作为正文段落
            elif current is not None:
                current["body"].append(text)
        
        elif t == "body":
            text = item["text"]
            # 判断是否是章节标题（短文本且包含特定关键词）
            is_title = len(text) < 30 and any(kw in text for kw in ["维度", "策略", "空间", "格局", "提示", "风险", "层面"])
            if is_title:
                # 作为新 section 的标题
                if current and (current["body"] or current["charts"]):
                    sections.append(current)
                current = {
                    "heading": text,
                    "body": [],
                    "charts": [],
                }
            elif current is not None:
                current["body"].append(text)
            else:
                # 没有标题的正文，创建一个默认 section
                current = {
                    "heading": "",
                    "body": [text],
                    "charts": [],
                }
        
        elif t == "chart_title":
            if current is not None:
                current["charts"].append({
                    "title": item["text"],
                    "image": None,
                    "data_source": "",
                })
        
        elif t == "image":
            if current is not None:
                caption = item.get("caption", "")
                if current["charts"] and not current["charts"][-1]["image"]:
                    current["charts"][-1]["image"] = item["src"]
                else:
                    current["charts"].append({
                        "title": caption,
                        "image": item["src"],
                        "data_source": "",
                    })
        
        elif t == "source":
            if current is not None and current["charts"]:
                src_text = item["text"].replace("资料来源：", "").replace("资料来源:", "").strip()
                if not current["charts"][-1]["data_source"]:
                    current["charts"][-1]["data_source"] = src_text
    
    if current and (current["body"] or current["charts"]):
        sections.append(current)
    
    # 4. 过滤掉分析师信息等非核心内容
    filtered = []
    skip_keywords = ["分析师", "邮箱", "执业编号", "电话", "研究所", "团队介绍", "研究员", "助理研究员", "销售", "地区", "北京", "上海", "深圳", "广州"]
    for s in sections:
        # 如果 section 的 body 中有超过一半是分析师信息，跳过
        skip_count = sum(1 for b in s["body"] if any(kw in b for kw in skip_keywords))
        if skip_count > len(s["body"]) * 0.3:
            continue
        # 移除分析师信息段落
        s["body"] = [b for b in s["body"] if not any(kw in b for kw in skip_keywords)]
        if s["body"] or s["charts"]:
            filtered.append(s)
    
    return filtered


def section_to_pages(section: dict) -> list[dict]:
    """把一个 section 拆分为多页
    
    规则：
    - 每页 1-2 张图
    - 每页最多 3 段正文
    - 图表标题在图片上方
    - 过滤掉图表备注（注：...）
    """
    heading = section["heading"]
    body = section["body"]
    charts = section["charts"]
    
    # 过滤掉图表备注
    body = [b for b in body if not b.startswith("注：") and not b.startswith("注:")]
    
    pages: list[dict] = []
    
    if not charts:
        # 纯文字页：按 3 段分页
        for i in range(0, max(1, len(body)), 3):
            chunk = body[i:i + 3]
            pages.append({
                "title": heading,
                "body": chunk,
                "images": [],
            })
        return pages
    
    # 有图的情况：按图分页，每 1-2 张图一页
    para_idx = 0
    chart_idx = 0
    
    while chart_idx < len(charts):
        # 取 1-2 张图
        page_charts = []
        for _ in range(2):
            if chart_idx < len(charts):
                c = charts[chart_idx]
                if c.get("image"):  # 只取有图的
                    page_charts.append(c)
                chart_idx += 1
        
        if not page_charts:
            continue
        
        # 给这页分配 2-3 段正文
        take = min(2, len(body) - para_idx)
        page_paras = body[para_idx:para_idx + take] if take > 0 else []
        para_idx += take
        
        pages.append({
            "title": heading,
            "body": page_paras,
            "images": [{"caption": c["title"], "src": c["image"], "data_source": c.get("data_source", "")} for c in page_charts],
        })
    
    # 剩余段落单独一页
    if para_idx < len(body):
        pages.append({
            "title": heading,
            "body": body[para_idx:],
            "images": [],
        })
    
    return pages


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--content", required=True)
    parser.add_argument("--media", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    
    with open(args.content, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    items = data.get("items", [])
    
    # 1. 解析出 sections
    sections = parse_content(items)
    
    # 2. 过滤空 section
    sections = [s for s in sections if s["body"] or s["charts"]]
    
    # 3. 每个 section 拆分为 pages
    all_pages: list[dict] = []
    for section in sections:
        pages = section_to_pages(section)
        all_pages.extend(pages)
    
    # 4. 统计
    total_images = sum(len(p.get("images", [])) for p in all_pages)
    total_paras = sum(len(p.get("body", [])) for p in all_pages)
    
    plan = {
        "total_pages": len(all_pages),
        "total_images": total_images,
        "total_paragraphs": total_paras,
        "slices": all_pages,
    }
    
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] {args.output}")
    print(f"     页数: {len(all_pages)}, 图片: {total_images}, 段落: {total_paras}")
    
    # 显示前几页预览
    for i, p in enumerate(all_pages[:8]):
        imgs = len(p.get("images", []))
        paras = len(p.get("body", []))
        print(f"     [{i+1:2d}] {p['title'][:40]:40s} | {paras}段 {imgs}图")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
