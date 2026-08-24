import argparse
import datetime as dt
import json
import re
import time
from dataclasses import dataclass

import pandas as pd
import requests


def _parse_dt(s: str) -> dt.datetime:
    # Example: "2026-04-28 00:00:00.000"
    return dt.datetime.strptime(s.split(".")[0], "%Y-%m-%d %H:%M:%S")


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
    escape = False
    for j in range(start, len(html)):
        ch = html[j]
        if in_str:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
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
    s = s.replace("\r", "")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def _pick_key_sentences(text: str, k: int = 3) -> list[str]:
    keywords = [
        "同比",
        "环比",
        "毛利",
        "净利",
        "营收",
        "销量",
        "产能",
        "产量",
        "成本",
        "费用",
        "现金流",
        "项目",
        "投产",
        "扩产",
        "开工",
        "订单",
        "需求",
        "价格",
        "景气",
        "指引",
        "预计",
        "展望",
        "2026",
    ]
    raw = re.split(r"(?<=[。；！?])\s*|\n+", text)
    sents: list[str] = []
    for s in raw:
        s = s.strip()
        if len(s) < 18 or len(s) > 220:
            continue
        sents.append(s)
    if not sents:
        return []

    def score(s: str) -> tuple[int, int]:
        kw = sum(1 for w in keywords if w in s)
        return (kw, len(s))

    ranked = sorted(sents, key=score, reverse=True)
    picked: list[str] = []
    seen = set()
    for s in ranked:
        if s in seen:
            continue
        key = re.sub(r"\s+", "", s)
        if key in seen:
            continue
        picked.append(s)
        seen.add(s)
        seen.add(key)
        if len(picked) >= k:
            break
    return picked


def _extract_risks(text: str) -> list[str]:
    m = re.search(r"风险提示[:：]\s*(.+)", text)
    if m:
        tail = m.group(1).split("\n")[0]
        items = re.split(r"[，、；;。]\s*", tail)
        items = [x.strip() for x in items if x.strip()]
        if items:
            return items[:4]
    return ["产品价格波动风险", "需求不及预期风险", "项目建设进度不及预期风险", "政策及环保安全风险"]


def _conclusion_from_text(text: str) -> str:
    if "超预期" in text:
        return "超预期"
    if "不及预期" in text or "低于预期" in text:
        return "低于预期"
    return "符合预期"


def _outlook_2026(text: str) -> str:
    candidates = []
    for s in re.split(r"(?<=[。；！?])\s*|\n+", text):
        s = s.strip()
        if 12 <= len(s) <= 80 and ("2026" in s or "展望" in s or "预计" in s):
            candidates.append(s)
    if candidates:
        return candidates[0]
    return "展望2026年，关注业绩修复与项目推进节奏。"


@dataclass
class Report:
    title: str
    publish_dt: dt.datetime
    info_code: str
    org_name: str


def _keep_report(r: Report) -> bool:
    if r.publish_dt < dt.datetime(2026, 1, 1):
        return False
    t = r.title
    annual_hit = ("点评" in t) and ("年报" in t or "年度" in t) and ("2025" in t or "25" in t)
    q1_hit = ("点评" in t) and ("一季报" in t or "一季度" in t)
    if ("2024" in t or "24" in t) and ("年报" in t or "年度" in t):
        annual_hit = False
    if ("2025" in t or "25" in t) and ("一季报" in t or "一季度" in t) and ("2026" not in t and "26" not in t):
        q1_hit = False
    return annual_hit or q1_hit


def _requests_session() -> requests.Session:
    s = requests.Session()
    s.trust_env = False  # ignore sandbox-injected proxies
    s.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept-Language": "zh-CN,zh;q=0.9",
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
            time.sleep(0.6 * (n + 1))
    raise RuntimeError(f"GET failed: {url}") from last_err


def _stockcode_6(stock_code: str) -> str:
    s = str(stock_code).strip()
    s = re.sub(r"\.(SH|SZ)$", "", s, flags=re.I)
    s = re.sub(r"[^0-9]", "", s)
    return s.zfill(6)


def _parse_reports_from_singlestock(html: str) -> list[Report]:
    obj_text = _extract_js_object(html, "initdata")
    data = json.loads(obj_text)
    reports: list[Report] = []
    for item in (data.get("data") or [])[:200]:
        try:
            reports.append(
                Report(
                    title=str(item.get("title", "")).strip(),
                    publish_dt=_parse_dt(str(item.get("publishDate", "")).strip()),
                    info_code=str(item.get("infoCode", "")).strip(),
                    org_name=str(item.get("orgSName") or item.get("orgName") or "").strip(),
                )
            )
        except Exception:
            continue
    return reports


def _parse_notice_content_from_zw(html: str) -> str:
    obj_text = _extract_js_object(html, "zwinfo")
    data = json.loads(obj_text)
    return str(data.get("notice_content") or "").strip()


def _summarize_reports(stock_name: str, reports: list[tuple[Report, str]]) -> dict:
    merged = "\n\n".join([_clean_text(txt) for _, txt in reports if txt.strip()])
    picked = _pick_key_sentences(merged, k=3)
    core_paras: list[str] = []
    for s in picked[:3]:
        core_paras.append(f"短核心句：{s}")

    outlook = _outlook_2026(merged)
    if core_paras:
        core_paras[-1] = core_paras[-1].rstrip("。") + f"。{outlook}"
    else:
        core_paras = [f"短核心句：{stock_name}在年报/一季报阶段的机构点评信息较少。{outlook}"]

    risks = _extract_risks(merged)
    conclusion = _conclusion_from_text(merged)

    excerpts = []
    for rep, txt in reports:
        short = _clean_text(txt)
        short = short[:600] + ("…" if len(short) > 600 else "")
        excerpts.append(
            {
                "publish_date": rep.publish_dt.strftime("%Y-%m-%d"),
                "org": rep.org_name,
                "title": rep.title,
                "infocode": rep.info_code,
                "excerpt": short,
            }
        )

    return {
        "core_viewpoints": "\n".join(core_paras),
        "risk_tips": "、".join(risks[:4]),
        "conclusion": conclusion[:4],
        "excerpts": json.dumps(excerpts, ensure_ascii=False),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--input",
        default=r"D:\JupyterFiles\huachuang\0528数据更新\0528可转债列表.xlsx",
        help="Excel path: 4 columns without header: bond_code,bond_name,stock_code,stock_name",
    )
    ap.add_argument(
        "--output",
        default=r"D:\JupyterFiles\huachuang\PY结果\0528_东方财富研报_2025年报与一季报_摘要汇总.xlsx",
        help="Output Excel path",
    )
    ap.add_argument(
        "--output-txt",
        default=r"D:\JupyterFiles\huachuang\PY结果\0528_东方财富研报_2025年报与一季报_摘抄与总结.txt",
        help="Output TXT path (formatted blocks for matched rows)",
    )
    ap.add_argument("--sleep", type=float, default=0.15, help="Sleep between requests")
    args = ap.parse_args()

    df = pd.read_excel(args.input, header=None)
    df.columns = ["bond_code", "bond_name", "stock_code", "stock_name"]
    df["stock_code_6"] = df["stock_code"].map(_stockcode_6)

    session = _requests_session()

    rows = []
    for _, row in df.iterrows():
        stock_code_6 = row["stock_code_6"]
        stock_name = str(row["stock_name"])
        sing_url = f"https://data.eastmoney.com/report/singlestock.jshtml?stockcode={stock_code_6}"
        try:
            html = _get_text(session, sing_url)
            reps = [r for r in _parse_reports_from_singlestock(html) if _keep_report(r)]
        except Exception:
            reps = []

        full_texts: list[tuple[Report, str]] = []
        for rep in reps[:6]:
            zw_url = f"https://data.eastmoney.com/report/zw_stock.jshtml?infocode={rep.info_code}"
            try:
                zw_html = _get_text(session, zw_url)
                notice = _parse_notice_content_from_zw(zw_html)
                if notice:
                    full_texts.append((rep, notice))
            except Exception:
                continue
            time.sleep(args.sleep)

        if full_texts:
            summ = _summarize_reports(stock_name, full_texts)
        else:
            summ = {"core_viewpoints": "", "risk_tips": "", "conclusion": "", "excerpts": "[]"}

        rows.append(
            {
                "bond_code": row["bond_code"],
                "bond_name": row["bond_name"],
                "stock_code": row["stock_code"],
                "stock_name": row["stock_name"],
                "matched_reports": len(full_texts),
                "核心观点": summ["core_viewpoints"],
                "风险提示": summ["risk_tips"],
                "结论": summ["conclusion"],
                "摘抄(含infocode)": summ["excerpts"],
            }
        )
        time.sleep(args.sleep)

    out = pd.DataFrame(rows)
    out.to_excel(args.output, index=False)

    # Use UTF-8 with BOM for better compatibility with Windows tools.
    with open(args.output_txt, "w", encoding="utf-8-sig") as f:
        for r in rows:
            core = (r.get("核心观点") or "").strip()
            if not core:
                continue
            f.write(f"{r['bond_code']} {r['bond_name']} / {r['stock_code']} {r['stock_name']}\n")
            f.write("【核心观点】\n")
            f.write(core + "\n")
            f.write("【风险提示】\n")
            f.write((r.get("风险提示") or "") + "\n")
            f.write("【结论】\n")
            f.write((r.get("结论") or "") + "\n\n")

    print(f"Wrote: {args.output}")
    print(f"Wrote: {args.output_txt}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
