import argparse
import datetime as dt
import json
import re
import time
from dataclasses import dataclass

import pandas as pd
import requests


def _requests_session() -> requests.Session:
    s = requests.Session()
    s.trust_env = False
    s.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Referer": "https://data.eastmoney.com/",
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
    if i >= len(html) or html[i] != "{":
        raise ValueError(f"Object for {var_name} not found")

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


def _parse_dt(s: str) -> dt.datetime:
    return dt.datetime.strptime(s.split(".")[0], "%Y-%m-%d %H:%M:%S")


def _stockcode_6(stock_code: str) -> str:
    s = str(stock_code).strip()
    s = re.sub(r"\.(SH|SZ)$", "", s, flags=re.I)
    s = re.sub(r"[^0-9]", "", s)
    return s.zfill(6)


@dataclass
class Report:
    title: str
    publish_dt: dt.datetime
    info_code: str
    org: str


def _parse_reports_from_singlestock(html: str) -> list[Report]:
    obj = _extract_js_object(html, "initdata")
    data = json.loads(obj)
    out: list[Report] = []
    for item in (data.get("data") or [])[:400]:
        try:
            out.append(
                Report(
                    title=str(item.get("title") or "").strip(),
                    publish_dt=_parse_dt(str(item.get("publishDate") or "").strip()),
                    info_code=str(item.get("infoCode") or "").strip(),
                    org=str(item.get("orgSName") or item.get("orgName") or "").strip(),
                )
            )
        except Exception:
            continue
    return out


def _fetch_notice_content(session: requests.Session, infocode: str) -> str:
    html = _get_text(session, f"https://data.eastmoney.com/report/zw_stock.jshtml?infocode={infocode}")
    obj = _extract_js_object(html, "zwinfo")
    data = json.loads(obj)
    return str(data.get("notice_content") or "").strip()


def _classify_by_content(text: str) -> set[str]:
    t = text or ""
    hits: set[str] = set()
    if re.search(r"2025\s*年\s*(年报|年度(报告)?)", t) and not re.search(
        r"(中报|半年报|三季报|Q2|Q3)", t, flags=re.I
    ):
        hits.add("annual_2025")
    if re.search(r"2026\s*年\s*(一季报|一季度)", t) or "2026Q1" in t or "26Q1" in t:
        hits.add("q1_2026")
    if ("年度报告" in t or "年报" in t) and "2025" in t:
        hits.add("annual_2025")
    if ("一季度" in t or "一季报" in t) and ("2026" in t or "26" in t):
        hits.add("q1_2026")
    return hits


def _extract_risks(text: str) -> list[str]:
    # Use unicode escapes to avoid console encoding issues.
    fengxian_tishi = "\u98ce\u9669\u63d0\u793a"  # 风险提示
    fengxian_yinsu = "\u98ce\u9669\u56e0\u7d20"  # 风险因素
    pat = rf"({fengxian_tishi}|{fengxian_yinsu})[:\uff1a]?\s*(.+?)(\n\s*\n|$)"
    m = re.search(pat, text or "", flags=re.S)
    if not m:
        return []
    risk = m.group(2).strip()
    # Strip following "投资建议" etc if same paragraph
    risk = re.split(r"\n\s*\u6295\u8d44\u5efa\u8bae[:\uff1a]", risk)[0].strip()
    items = [x.strip() for x in re.split(r"[、，；;。]\s*", risk) if x.strip()]
    return items[:6]


def _extract_risks_from_research_fallback(text: str) -> list[str]:
    # If no explicit "风险提示" section, try to mine short "...风险" phrases from the whole text.
    t = (text or "").replace("\r", "")
    candidates = re.findall(r"[\u4e00-\u9fffA-Za-z0-9]{2,30}?风险", t)
    stop = {"重大风险", "风险", "可能面对的风险", "可能存在的风险"}
    out: list[str] = []
    seen = set()
    for x in candidates:
        x = re.sub(r"\s+", "", x)
        if len(x) > 24 or x in stop:
            continue
        if any(p in x for p in ["详细描述", "可能面对", "可能存在", "公司已在本报告中"]):
            continue
        if x in seen:
            continue
        seen.add(x)
        out.append(x)
        if len(out) >= 4:
            break
    return out


def _is_rev_np_yoy_sentence(s: str) -> bool:
    s0 = s.replace(" ", "")
    has_rev = any(k in s0 for k in ["营收", "营业收入", "收入"])
    has_np = any(k in s0 for k in ["归母净利润", "归母净利", "净利润"])
    has_yoy = any(k in s0 for k in ["同比", "环比"])
    # Only filter the repetitive headline-type sentences about revenue / net profit YoY/QoQ.
    return (has_rev or has_np) and has_yoy and ("毛利" not in s0) and ("费用" not in s0) and ("现金流" not in s0)


def _summarize_paragraph(text: str) -> str:
    text = (text or "").replace("\r", "").strip()
    sentences = [s.strip() for s in re.split(r"[。；！？?]\s*|\n+", text) if s.strip()]
    keywords = [
        "营收",
        "收入",
        "归母",
        "净利润",
        "毛利",
        "费用",
        "现金流",
        "订单",
        "客户",
        "产能",
        "产量",
        "价格",
        "成本",
        "指引",
        "预计",
        "展望",
        "同比",
        "环比",
        "投产",
        "产线",
        "产品",
    ]
    candidates = [s for s in sentences if 16 <= len(s) <= 160 and not _is_rev_np_yoy_sentence(s)]
    if not candidates:
        return "研报正文可用信息较少。"

    def score(s: str) -> tuple[int, int]:
        return (sum(1 for w in keywords if w in s), len(s))

    ranked = sorted(candidates, key=score, reverse=True)
    picked: list[str] = []
    for s in ranked:
        if s in picked:
            continue
        picked.append(s)
        if len(picked) >= 5:
            break
    return "；".join(picked[:3]) + ("。" if not picked[-1].endswith(("。", "！", "？")) else "")


def _get_jsonp(session: requests.Session, url: str, params: dict) -> dict:
    params = dict(params)
    params.setdefault("cb", "cb")
    last_err: Exception | None = None
    for n in range(6):
        try:
            r = session.get(url, params=params, timeout=30)
            r.raise_for_status()
            raw = r.text
            break
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(0.7 * (n + 1))
    else:
        raise RuntimeError("JSONP request failed") from last_err
    m = re.match(r"cb\((.*)\)\s*$", raw, flags=re.S)
    if not m:
        raise ValueError("Invalid JSONP response")
    return json.loads(m.group(1))


def _fetch_annual_report_notice_content(session: requests.Session, stock_code_6: str) -> str | None:
    base = "https://np-anotice-stock.eastmoney.com/api/security/ann"
    params = {
        "ann_type": "A",
        "client_source": "web",
        "stock_list": stock_code_6,
        "page_index": 1,
        "page_size": 200,
        "sr": -1,
        "st": "notice_date",
    }
    try:
        items = session.get(base, params=params, timeout=30).json()["data"]["list"]
    except Exception:
        return None

    annual_kw = "\u5e74\u5ea6\u62a5\u544a"  # 年度报告
    annual_2025 = "\u0032\u0030\u0032\u0035\u5e74\u5e74\u5ea6\u62a5\u544a"  # 2025年年度报告
    art_code = None
    for it in items:
        title = it.get("title") or ""
        if annual_2025 in title and title.endswith(annual_kw):
            art_code = it.get("art_code")
            break
    if not art_code:
        return None

    api = "https://np-cnotice-stock.eastmoney.com/api/content/ann"
    pages: list[str] = []
    seen = set()
    # The annual report content is paginated; page 1 is usually front matter.
    # Fetch multiple pages until we have enough to parse "分产品" + margin change reasons + risks.
    for p in range(1, 16):
        try:
            payload = _get_jsonp(session, api, {"art_code": art_code, "client_source": "web", "page_index": p})
            content = str((payload.get("data") or {}).get("notice_content") or "").strip()
        except Exception:
            time.sleep(0.4)
            continue
        if not content:
            break
        key = content[:300]
        if key in seen:
            break
        seen.add(key)
        pages.append(content)
        joined = "\n".join(pages)
        has_products = "\u5206\u4ea7\u54c1" in joined
        has_fin = ("\u8425\u4e1a\u6536\u5165" in joined) and ("\u6bdb\u5229\u7387" in joined)
        has_reason = ("\u4e0b\u964d\u4e3b\u8981\u539f\u56e0\u662f" in joined) or ("\u4e3b\u8981\u539f\u56e0\u662f" in joined)
        has_risk_anchor = ("\u98ce\u9669\u63d0\u793a" in joined) or ("\u91cd\u5927\u98ce\u9669\u63d0\u793a" in joined) or ("\u98ce\u9669\u53ca\u5e94\u5bf9\u63aa\u65bd" in joined)
        risk_phrase_hits = joined.count("\u98ce\u9669")
        if has_products and has_fin and has_reason and (has_risk_anchor or risk_phrase_hits >= 5):
            break
    joined = "\n".join(pages).strip()
    return joined or None


def _parse_income_cost_table(annual_text: str) -> list[dict]:
    t = (annual_text or "").replace("\r", "")
    anchor = "\u5360\u516c\u53f8\u8425\u4e1a\u6536\u5165\u6216\u8425\u4e1a\u5229\u6da6"  # 占公司营业收入或营业利润
    pos = t.find(anchor)
    if pos == -1:
        pos = t.find("\u8425\u4e1a\u6536\u5165\u6784\u6210")  # 营业收入构成
    if pos == -1:
        return []
    chunk = t[pos : pos + 9000]

    def to_num(x: str) -> float:
        x = x.replace(",", "").strip()
        try:
            return float(x)
        except Exception:
            return float("nan")

    rows: list[dict] = []
    current_section = None
    for line in chunk.splitlines():
        line = re.sub(r"\s+", " ", line).strip()
        if not line:
            continue
        if line in ("\u5206\u4ea7\u54c1", "\u5206\u884c\u4e1a", "\u5206\u5730\u533a", "\u5206\u9500\u552e\u6a21\u5f0f"):
            current_section = line
            continue
        if current_section != "\u5206\u4ea7\u54c1":
            continue

        m = re.match(
            r"^(?P<name>.+?) (?P<rev>[0-9,]+\.[0-9]+) (?P<cost>[0-9,]+\.[0-9]+) (?P<gm>[0-9]+\.[0-9]+)% .+? (?P<gmchg>-?[0-9]+\.[0-9]+)%$",
            line,
        )
        if not m:
            continue
        name = m.group("name").strip()
        rev = to_num(m.group("rev"))
        cost = to_num(m.group("cost"))
        gm = float(m.group("gm")) / 100.0
        gmchg = float(m.group("gmchg"))
        rows.append({"name": name, "rev": rev, "cost": cost, "gp": rev - cost, "gm": gm, "gmchg_pct": gmchg})
    return rows


def _extract_gm_change_reasons(annual_text: str) -> str:
    t = (annual_text or "").replace("\r", "")
    pos = t.find("\u4e0b\u964d\u4e3b\u8981\u539f\u56e0\u662f")  # 下降主要原因是
    if pos == -1:
        pos = t.find("\u4e3b\u8981\u539f\u56e0\u662f")  # 主要原因是
    if pos == -1:
        return ""
    chunk = t[pos : pos + 900]
    sents = [s.strip() for s in re.split(r"[。；！？?]\s*|\n+", chunk) if s.strip()]
    out: list[str] = []
    for s in sents:
        if any(k in s for k in ["受", "影响", "导致", "叠加", "固定成本", "竞争", "定价", "回款", "付款", "产能"]):
            out.append(s)
        if len(out) >= 3:
            break
    return "；".join(out)


def _extract_risks_from_annual(annual_text: str) -> list[str]:
    t = (annual_text or "").replace("\r", "")
    stop_phrases = [
        "\u91cd\u5927\u98ce\u9669\u63d0\u793a",  # 重大风险提示
        "\u98ce\u9669\u63d0\u793a",  # 风险提示
        "\u516c\u53f8\u5df2\u5728\u672c\u62a5\u544a\u4e2d\u8be6\u7ec6\u63cf\u8ff0",  # 公司已在本报告中详细描述
        "\u8be6\u7ec6\u63cf\u8ff0",  # 详细描述
        "\u53ef\u80fd\u5b58\u5728",  # 可能存在
        "\u53ef\u80fd\u9762\u5bf9",  # 可能面对
    ]

    def is_generic(x: str) -> bool:
        return any(p in x for p in stop_phrases) or len(x) <= 3 or x in ("重大风险", "风险", "可能面对的风险")

    # 1) Try to locate a "风险提示/重大风险提示" neighborhood and extract explicit "...风险" phrases.
    pos = t.find("\u91cd\u5927\u98ce\u9669\u63d0\u793a")
    if pos == -1:
        pos = t.find("\u98ce\u9669\u63d0\u793a")
    chunks: list[str] = []
    if pos != -1:
        chunks.append(t[pos : pos + 2200])

    # 2) Also try broader risk section in MDA, typically "风险及应对措施/公司未来发展的展望".
    for anchor in [
        "\u98ce\u9669\u53ca\u5e94\u5bf9\u63aa\u65bd",  # 风险及应对措施
        "\u516c\u53f8\u7ecf\u8425\u4e2d\u53ef\u80fd\u9762\u4e34\u7684\u98ce\u9669",  # 公司经营中可能面临的风险
        "\u516c\u53f8\u672a\u6765\u53d1\u5c55\u7684\u5c55\u671b",  # 公司未来发展的展望
        "\u98ce\u9669\u56e0\u7d20",  # 风险因素
    ]:
        p = t.find(anchor)
        if p != -1:
            chunks.append(t[p : p + 8000])

    # Collect candidates
    candidates: list[str] = []
    for chunk in chunks:
        # numbered risks like （1）xxx风险 / 1、xxx风险
        candidates += re.findall(r"[（(]?\d+[）)]\s*([^\n。；]{3,40}?风险)", chunk)
        candidates += re.findall(r"\d+、\s*([^\n。；]{3,40}?风险)", chunk)
        candidates += re.findall(r"[\u4e00-\u9fffA-Za-z0-9]{2,30}?风险", chunk)

    # Clean and filter
    cleaned: list[str] = []
    for x in candidates:
        x = re.sub(r"\s+", "", x)
        x = x.strip("：:，,；;。.")
        if not x.endswith("风险"):
            continue
        # drop too-long generic sentences
        if len(x) > 24:
            continue
        if is_generic(x):
            continue
        cleaned.append(x)

    # De-dup preserve order and cap 4
    seen = set()
    out: list[str] = []
    for x in cleaned:
        if x in seen:
            continue
        seen.add(x)
        out.append(x)
        if len(out) >= 4:
            break
    return out


def _annual_fallback_core(annual_text: str) -> str:
    rows = _parse_income_cost_table(annual_text)
    reason = _extract_gm_change_reasons(annual_text)
    parts: list[str] = []
    if rows:
        rows = sorted(rows, key=lambda x: (x["rev"] if x["rev"] == x["rev"] else -1), reverse=True)
        show = rows[:2]
        segs = []
        for r in show:
            segs.append(
                f"{r['name']}：收入{r['rev']/1e8:.2f}亿元、毛利{r['gp']/1e8:.2f}亿元、毛利率{r['gm']*100:.2f}%（同比{r['gmchg_pct']:+.2f}pct）"
            )
        parts.append("分产品看，" + "；".join(segs))
    if reason:
        reason = reason.replace("下降主要原因是", "").replace("主要原因是", "").lstrip("：: ")
        parts.append("毛利率变动主要受" + reason)
    return "；".join(parts).strip()


def _infer_risks_from_annual_text(annual_text: str) -> list[str]:
    t = (annual_text or "").replace("\r", "")
    risks: list[str] = []

    def add(x: str) -> None:
        if x not in risks:
            risks.append(x)

    # Map typical annual-report wording to explicit risks
    if any(k in t for k in ["工程款", "回款", "付款进度", "结算", "应收账款"]):
        add("回款及应收账款回收风险")
    if any(k in t for k in ["竞争加剧", "价格战", "定价下行", "项目定价"]):
        add("行业竞争加剧及项目定价下行风险")
    if any(k in t for k in ["产能", "产量不饱和", "固定成本分摊", "产能过剩"]):
        add("产能利用率不足及固定成本分摊压力风险")
    if any(k in t for k in ["需求放缓", "市场需求", "景气度下滑", "销量不及预期"]):
        add("需求不及预期风险")
    if any(k in t for k in ["原材料", "涨价", "价格波动"]):
        add("原材料价格波动风险")

    return risks[:4]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--input",
        default=r"D:\JupyterFiles\huachuang\PY结果\0528可转债列表.xlsx",
        help="0528 list path (either 4-cols no-header, or the PY结果 version with headers)",
    )
    ap.add_argument(
        "--output",
        default=r"D:\JupyterFiles\huachuang\PY结果\0528_研报核心观点_风险提示.xlsx",
        help="Output xlsx",
    )
    ap.add_argument("--max-scan-reports", type=int, default=20, help="Max recent reports to open per stock")
    ap.add_argument("--sleep", type=float, default=0.15)
    args = ap.parse_args()

    raw = pd.read_excel(args.input)
    has_header = {"转债代码", "转债简称", "正股代码", "正股简称"}.issubset(set(raw.columns))
    if has_header:
        bond_code_col = "转债代码"
        bond_name_col = "转债简称"
        stock_code_col = "正股代码"
        stock_name_col = "正股简称"
        df = raw.copy()
    else:
        df = pd.read_excel(args.input, header=None)
        df.columns = ["bond_code", "bond_name", "stock_code", "stock_name"]
        bond_code_col = "bond_code"
        bond_name_col = "bond_name"
        stock_code_col = "stock_code"
        stock_name_col = "stock_name"
    df["stock_code_6"] = df[stock_code_col].map(_stockcode_6)

    session = _requests_session()
    total = len(df)
    core_col = "核心观点"
    risk_col = "风险提示"
    df[core_col] = ""
    df[risk_col] = ""

    for i in range(total):
        stock_code_6 = str(df.at[i, "stock_code_6"])
        stock_name = str(df.at[i, stock_name_col])
        print(f"[{i+1}/{total}] {stock_code_6} {stock_name} ...", flush=True)

        core_parts: list[str] = []
        risks: list[str] = []

        try:
            html = _get_text(session, f"https://data.eastmoney.com/report/singlestock.jshtml?stockcode={stock_code_6}")
            reports = _parse_reports_from_singlestock(html)
            reports = sorted(
                [r for r in reports if r.publish_dt >= dt.datetime(2026, 1, 1)],
                key=lambda x: x.publish_dt,
                reverse=True,
            )
        except Exception:
            reports = []

        annual_text: str | None = None
        q1_text: str | None = None
        annual_notice_text: str | None = None

        for rep in reports[: args.max_scan_reports]:
            if annual_text and q1_text:
                break
            try:
                text = _fetch_notice_content(session, rep.info_code)
                kinds = _classify_by_content(text)
                if (not annual_text) and ("annual_2025" in kinds):
                    annual_text = text
                if (not q1_text) and ("q1_2026" in kinds):
                    q1_text = text
            except Exception:
                continue
            finally:
                time.sleep(args.sleep)

        annual_summary = _summarize_paragraph(annual_text) if annual_text else ""
        q1_summary = _summarize_paragraph(q1_text) if q1_text else ""
        # If no research report matched, fallback to 2025 annual report (公告) for core+risks.
        if not annual_summary and not q1_summary:
            annual_notice_text = _fetch_annual_report_notice_content(session, stock_code_6)
            if annual_notice_text:
                annual_summary = _annual_fallback_core(annual_notice_text)
                if not risks:
                    risks.extend(_extract_risks_from_annual(annual_notice_text))
                if not risks:
                    risks.extend(_infer_risks_from_annual_text(annual_notice_text))
        if annual_summary and q1_summary:
            core = f"【核心观点】年报方面，{annual_summary}；一季报方面，{q1_summary}"
        elif annual_summary:
            core = f"【核心观点】{annual_summary}"
        elif q1_summary:
            core = f"【核心观点】{q1_summary}"
        else:
            core = "【核心观点】未抓取到研报，且年报关键分项信息提取失败。"

        if core:
            core_parts = [core]

        if annual_text:
            risks.extend(_extract_risks(annual_text))
            if not risks:
                risks.extend(_extract_risks_from_research_fallback(annual_text))
        if q1_text:
            risks.extend(_extract_risks(q1_text))
            if not risks:
                risks.extend(_extract_risks_from_research_fallback(q1_text))

        risk_items = [x for x in dict.fromkeys([r for r in risks if r])]
        risk_text = ("【风险提示】" + "、".join(risk_items)) if risk_items else ""

        df.at[i, core_col] = "\n".join(core_parts).strip()
        df.at[i, risk_col] = risk_text.strip()
        time.sleep(args.sleep)

    # drop helper
    df = df.drop(columns=["stock_code_6"])
    try:
        df.to_excel(args.output, index=False)
        print(f"Wrote: {args.output}")
    except PermissionError:
        alt = re.sub(r"\.xlsx$", "_v6.xlsx", args.output, flags=re.I)
        df.to_excel(alt, index=False)
        print(f"Wrote (fallback): {alt}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
