import argparse
import datetime as dt
import json
import re
import time
from dataclasses import dataclass

import requests


def _requests_session() -> requests.Session:
    s = requests.Session()
    s.trust_env = False
    s.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Referer": "https://guba.eastmoney.com/",
        }
    )
    return s


def _get_text(session: requests.Session, url: str, tries: int = 3) -> str:
    last_err: Exception | None = None
    for n in range(tries):
        try:
            r = session.get(url, timeout=30)
            r.raise_for_status()
            r.encoding = r.apparent_encoding or "utf-8"
            return r.text
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(0.7 * (n + 1))
    raise RuntimeError(f"GET failed: {url}") from last_err


def _extract_js_object(html: str, var_name: str) -> str:
    m = re.search(rf"var\s+{re.escape(var_name)}\s*=\s*", html)
    if not m:
        raise ValueError(f"Variable not found: {var_name}")
    i = m.end()
    while i < len(html) and html[i] != "{":
        i += 1
    start = i
    depth = 0
    in_str = False
    esc = False
    for j in range(start, len(html)):
        ch = html[j]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return html[start : j + 1]
    raise ValueError(f"Unterminated object for {var_name}")


def _clean_text(s: str) -> str:
    s = (s or "").replace("\r", "").strip()
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s


@dataclass
class GubaItem:
    title: str
    url: str
    display_time: str  # "05-19 01:19" or "2026-05-19 01:19"


def _parse_list_items(list_html: str) -> list[GubaItem]:
    items: list[GubaItem] = []
    # Prefer rows that look like "研报/资讯" items: anchors with data-posttype="2".
    for tr in re.findall(r'<tr class="listitem".*?</tr>', list_html, flags=re.S):
        if 'data-posttype="2"' not in tr:
            continue
        href_m = re.search(r'href="(?P<href>/news,[^"]+?\.html)"', tr)
        if not href_m:
            continue
        href = href_m.group("href")
        title_m = re.search(rf'href="{re.escape(href)}"\s*>(?P<t>.*?)</a>', tr, flags=re.S)
        title = re.sub(r"<.*?>", "", title_m.group("t") if title_m else "").strip()
        time_m = re.search(r'<div class="update[^"]*">\s*(?P<t>[^<]+)\s*</div>', tr)
        display_time = (time_m.group("t").strip() if time_m else "").strip()
        items.append(GubaItem(title=title, url="https://guba.eastmoney.com" + href, display_time=display_time))
    seen: set[str] = set()
    ordered: list[GubaItem] = []
    for it in items:
        if it.url in seen:
            continue
        seen.add(it.url)
        ordered.append(it)
    return ordered


def _extract_source_id_from_news(news_html: str) -> str | None:
    # Page embeds JSON with "post_source_id":"AP...."
    m = re.search(r'"post_source_id"\s*:\s*"(?P<id>AP[0-9]+)"', news_html)
    if m:
        return m.group("id")
    m = re.search(r'"post_pdf_url"\s*:\s*"https://pdf\.dfcfw\.com/pdf/H3_(?P<id>AP[0-9]+)_1\.pdf', news_html)
    if m:
        return m.group("id")
    return None


def _extract_publish_time_from_news(news_html: str) -> dt.datetime | None:
    m = re.search(r'"post_publish_time"\s*:\s*"(?P<t>\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})"', news_html)
    if not m:
        return None
    try:
        return dt.datetime.strptime(m.group("t"), "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def _fetch_report_text_by_infocode(session: requests.Session, infocode: str) -> str:
    url = f"https://data.eastmoney.com/report/zw_stock.jshtml?infocode={infocode}"
    html = _get_text(session, url)
    obj = _extract_js_object(html, "zwinfo")
    data = json.loads(obj)
    return str(data.get("notice_content") or "").strip()


def _is_2025_annual(title: str) -> bool:
    return ("点评" in title) and ("2025" in title or "25" in title) and ("年报" in title or "年度" in title)


def _is_2026_q1(title: str) -> bool:
    return ("点评" in title) and ("一季报" in title or "一季度" in title) and ("2026" in title or "26" in title)


def _classify_by_content(text: str) -> set[str]:
    t = text or ""
    hits: set[str] = set()
    annual_pat = re.search(r"2025\s*年\s*(年报|年度(报告)?)", t)
    if annual_pat and not re.search(r"(中报|半年报|三季报|Q2|Q3)", t, flags=re.I):
        hits.add("annual_2025")
    if re.search(r"2026\s*年\s*(一季报|一季度)\b", t) or "2026Q1" in t or "26Q1" in t:
        hits.add("q1_2026")
    # Some reports might omit the year in正文 but have "2025年" near "年度报告"
    if ("年度报告" in t or "年报" in t) and "2025" in t:
        hits.add("annual_2025")
    if ("一季度" in t or "一季报" in t) and ("2026" in t or "26" in t):
        hits.add("q1_2026")
    return hits


def _summarize_as_paragraph(text: str, kind: str) -> str:
    # Simple extractive summarization: choose 4-6 key sentences.
    kw = [
        "营收",
        "归母",
        "净利润",
        "毛利",
        "费用",
        "现金流",
        "订单",
        "产能",
        "产量",
        "价格",
        "成本",
        "指引",
        "预计",
        "展望",
        "新品",
        "客户",
    ]
    sentences = [s.strip() for s in re.split(r"(?<=[。；！?])\s*|\n+", _clean_text(text)) if s.strip()]
    candidates = [s for s in sentences if 18 <= len(s) <= 220]
    if not candidates:
        return f"{kind}：正文可用信息较少。"

    def score(s: str) -> tuple[int, int]:
        return (sum(1 for w in kw if w in s), len(s))

    ranked = sorted(candidates, key=score, reverse=True)
    picked: list[str] = []
    for s in ranked:
        if s in picked:
            continue
        picked.append(s)
        if len(picked) >= 6:
            break
    return f"{kind}：" + " ".join(picked)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--stock", default="688533", help="Stock code, e.g. 688533")
    ap.add_argument("--pages", type=int, default=5, help="How many list pages to scan")
    ap.add_argument("--scan-limit", type=int, default=80, help="Max news items to open and classify")
    ap.add_argument("--sleep", type=float, default=0.2)
    args = ap.parse_args()

    session = _requests_session()

    items: list[GubaItem] = []
    for p in range(1, args.pages + 1):
        url = f"https://guba.eastmoney.com/list,{args.stock},2,f_{p}.html" if p > 1 else f"https://guba.eastmoney.com/list,{args.stock},2,f.html"
        html = _get_text(session, url)
        items.extend(_parse_list_items(html))
        time.sleep(args.sleep)

    # Titles on guba report list are often a short headline and may not include
    # "2025年报/2026一季报". Classify by正文内容 instead.
    annual: tuple[GubaItem, str, str] | None = None
    q1: tuple[GubaItem, str, str] | None = None

    # Prioritize potentially relevant headlines first to reduce remote calls.
    key_words = ["年报", "年度", "一季", "一季度", "季度", "Q1", "25年", "26年", "2025", "2026"]
    priority = [it for it in items if any(w in it.title for w in key_words)]
    rest = [it for it in items if it not in priority]
    scan_items = priority + rest

    scanned = 0
    for it in scan_items:
        if annual and q1:
            break
        try:
            news_html = _get_text(session, it.url)
            pub_dt = _extract_publish_time_from_news(news_html)
            if pub_dt and pub_dt < dt.datetime(2026, 1, 1):
                continue
            infocode = _extract_source_id_from_news(news_html)
            if not infocode:
                continue
            report_text = _fetch_report_text_by_infocode(session, infocode)
            if not report_text:
                continue
            scanned += 1
            kinds = _classify_by_content(report_text)
            if (not annual) and ("annual_2025" in kinds):
                annual = (it, infocode, report_text)
            if (not q1) and ("q1_2026" in kinds):
                q1 = (it, infocode, report_text)
        except Exception:
            continue
        finally:
            time.sleep(args.sleep)
        if scanned >= args.scan_limit:
            break

    # Console output in required structure for manual confirmation
    if annual:
        it, infocode, text = annual
        print(f"【2025年报点评】{it.title}（infocode={infocode}，列表时间={it.display_time}）")
        print(_summarize_as_paragraph(text, "年报总结"))
        print()
    else:
        print("【2025年报点评】未在扫描页内找到匹配条目")
        print()

    if q1:
        it, infocode, text = q1
        print(f"【2026年一季报点评】{it.title}（infocode={infocode}，列表时间={it.display_time}）")
        print(_summarize_as_paragraph(text, "一季报总结"))
        print()
    else:
        print("【2026年一季报点评】未在扫描页内找到匹配条目")
        print()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
