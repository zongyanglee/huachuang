# -*- coding: utf-8 -*-
"""生成未来 5 个交易日的可转债发行/上市日历（PNG + Excel）。

本脚本独立复写日度数据更新中的 iFinD 取数逻辑，不 import 原更新脚本。
默认通过 THS_Date_Offset 从运行日向后偏移 5 个交易日。接口返回运行日及
未来 5 个交易日，共 6 个日期；日历表头排除运行日，仅展示未来 5 个交易日：

* 发行前一交易日：原股东股权登记日
* 发行日：原股东配售日、网上发行日
* 发行后第一交易日：中签率
* 发行后第二交易日：中签结果
* 上市日：上市

运行示例：
    py 发行日历.py
    py 发行日历.py --date 2026-08-02
    py 发行日历.py --input-json 已整理的数据.json
"""

from __future__ import annotations

import argparse
from configparser import ConfigParser
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
from contextlib import redirect_stdout
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
with redirect_stdout(io.StringIO()):
    from iFinDPy import (
        THS_BD,
        THS_DR,
        THS_DataStatistics,
        THS_Date_Offset,
        THS_GetErrorInfo,
        THS_iFinDLogin,
    )


SCRIPT_DIR = Path(__file__).resolve().parents[2]
IMAGE_BUILDER = SCRIPT_DIR / "scripts" / "build_convertible_bond_calendar_image.mjs"
BUNDLED_NODE = Path(
    os.environ.get(
        "CODEX_BUNDLED_NODE",
        r"C:\Users\micub\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe",
    )
)

IFIND_CREDENTIAL_FILE = SCRIPT_DIR / "private/ifind账号.txt"
THS_LOGIN_OK_CODES = (0, -201)


def load_ifind_credentials() -> tuple[str, str]:
    """从项目目录的 ifind账号.txt 读取统一登录账号。"""
    if not IFIND_CREDENTIAL_FILE.is_file():
        raise FileNotFoundError(f"未找到iFinD账号文件：{IFIND_CREDENTIAL_FILE}")
    config = ConfigParser(interpolation=None)
    config.read(IFIND_CREDENTIAL_FILE, encoding="utf-8")
    username = config.get("ifind", "username", fallback="").strip()
    password = config.get("ifind", "password", fallback="").strip()
    if not username or not password:
        raise RuntimeError("ifind账号.txt中的[ifind] username或password为空")
    return username, password


def print_ifind_usage() -> None:
    """显示iFinD各数据项的已用额度比例。"""
    try:
        result = THS_DataStatistics()
        tables = result.get("tables", {}) if isinstance(result, dict) else {}
        if not tables:
            detail = result.get("errmsg", "未返回额度数据") if isinstance(result, dict) else str(result)
            print(f"[警告] iFinD使用额度查询失败：{detail}")
            return
        print("iFinD使用额度：")
        for key, value in tables.items():
            ratio = value.get("ratio", "N/A") if isinstance(value, dict) else value
            print(f"{key} 已用：{ratio}")
    except Exception as exc:
        print(f"[警告] iFinD使用额度查询失败：{exc}")

TITLE = "【华创固收 | 周冠南团队】可转债发行日历"
SUBTITLE = "周冠南 13488825234/张文星19121321217/李宗阳16619960676"
RATING_AGENCY_ABBREVIATIONS = {
    "中诚信国际信用评级有限责任公司": "中诚信",
    "东方金诚国际信用评估有限公司": "东方金诚",
    "联合资信评估股份有限公司": "联合资信",
    "上海新世纪资信评估投资服务有限公司": "上海新世纪",
    "中证鹏元资信评估股份有限公司": "中证鹏元",
    "大公国际资信评估有限公司": "大公国际",
}
COLORS = {
    "上市": "#C00000",
    "原股东股权登记日": "#FCE4D6",
    "原股东配售日，网上发行日": "#FFB7B7",
    "中签率": "#E7E6E6",
    "中签结果": "#D9E1F2",
}

EVENT_FIELDS = (
    "ths_convertible_debt_short_name_cbond;"
    "ths_online_issue_date_cbond;"
    "ths_listed_date_cbond"
)
EVENT_COLUMNS = [
    "简称",
    "发行日期",
    "上市日期",
]

DISPLAY_FIELDS = (
    "ths_online_issue_pur_code_cbond;"
    "ths_object_the_sw_bond;"
    "ths_issue_total_amt_bond;"
    "ths_issuer_entrust_rating_org_bond;"
    "ths_debt_rating_primary_rating_agency_bond"
)
DISPLAY_COLUMNS = [
    "网上申购代码",
    "所属行业",
    "发行规模",
    "评级公司",
    "债项评级",
]

ISSUING_ISSUE_FIELDS = (
    "p04647_f001:Y,p04647_f002:Y,p04647_f004:Y,p04647_f009:Y,"
    "p04647_f026:Y,p04647_f042:Y,p04647_f043:Y"
)

PENDING_ISSUE_FIELDS = (
    "p04649_f001:Y,p04649_f002:Y,p04649_f004:Y,"
    "p04649_f009:Y,p04649_f026:Y,p04649_f043:Y,p04649_f044:Y"
)

UNLISTED_ISSUE_FIELDS = (
    "jydm:Y,jydm_mc:Y,p05479_f001:Y,p05479_f044:Y,"
    "p05479_f046:Y,p05479_f019:Y"
)


def parse_ifind_date(value: Any) -> date | None:
    if value is None or (not isinstance(value, (date, datetime)) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    for fmt in ("%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    parsed = pd.to_datetime(text, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


def clean_scalar(value: Any) -> Any:
    if value is None or (not isinstance(value, (date, datetime)) and pd.isna(value)):
        return None
    if hasattr(value, "item"):
        value = value.item()
    return value


def login_ifind(ths_id: str | None = None, ths_password: str | None = None) -> None:
    if not ths_id or not ths_password:
        file_id, file_password = load_ifind_credentials()
        ths_id = ths_id or file_id
        ths_password = ths_password or file_password
    with redirect_stdout(io.StringIO()):
        code = THS_iFinDLogin(ths_id, ths_password)
    if code not in THS_LOGIN_OK_CODES:
        try:
            info = THS_GetErrorInfo(code)
            message = info.get("errmsg", str(info)) if isinstance(info, dict) else str(info)
        except Exception:
            message = "无法取得详细错误信息"
        raise RuntimeError(f"iFinD 登录失败（状态码 {code}）：{message}")
    print_ifind_usage()


def result_frame(result: Any, context: str, *, allow_empty: bool = False) -> pd.DataFrame:
    frame = getattr(result, "data", None)
    if isinstance(frame, pd.DataFrame) and (allow_empty or not frame.empty):
        return frame.copy()
    code = getattr(result, "errorcode", "未知")
    message = getattr(result, "errmsg", "无返回信息")
    # iFinD 的报表在没有符合条件的记录时，可能返回 -4001 / "no data"。
    # 对明确允许为空的列表查询，将其视为正常空表而不是中断整个日历。
    if allow_empty and (code in (0, -4001) or str(message).strip().lower() == "no data."):
        return pd.DataFrame()
    raise RuntimeError(f"{context}返回空数据（状态码 {code}）：{message}")


def fetch_p05479_unlisted_codes(as_of: date) -> list[str]:
    """从 p05479 获取已发行未上市转债，并剔除 NQ 与“定转”品种。"""
    result = THS_DR(
        "p05479",
        (
            "jyzt=2;sfdb=1;jysc=1;sszt=213006;"
            f"edate={as_of.strftime('%Y%m%d')};gnfl=0"
        ),
        UNLISTED_ISSUE_FIELDS,
        "format:dataframe",
    )
    frame = result_frame(result, "已发行未上市转债 p05479", allow_empty=True)
    if frame.empty:
        return []
    required = {"jydm", "jydm_mc"}
    if not required.issubset(frame.columns):
        raise RuntimeError(f"p05479 缺少字段 {sorted(required - set(frame.columns))}：{list(frame.columns)}")

    raw_codes = frame["jydm"].astype(str).str.strip()
    short_names = frame["jydm_mc"].astype(str).str.strip()
    keep = ~raw_codes.str.contains(".NQ", case=False, regex=False, na=False)
    keep &= ~short_names.str.contains("定转", regex=False, na=False)

    codes = raw_codes[keep].map(ensure_market_suffix)
    codes = codes[
        codes.str.match(
            r"^(?:110|111|113|118|123|127|128)\d{3}\.(?:SH|SZ)$",
            na=False,
        )
    ]
    return list(dict.fromkeys(codes.tolist()))


def fetch_issue_report(report_id: str, fields: str, label: str) -> pd.DataFrame:
    """获取发行阶段报表，并把各报表的 f001 统一为转债代码。"""
    result = THS_DR(
        report_id,
        "zqlx=640007;gnfl=0",
        fields,
        "format:dataframe",
    )
    frame = result_frame(result, f"{label} {report_id}", allow_empty=True)
    if frame.empty:
        return pd.DataFrame(columns=["转债代码"])
    code_field = f"{report_id}_f001"
    if code_field not in frame.columns:
        raise RuntimeError(f"{report_id} 缺少转债代码字段 {code_field}：{list(frame.columns)}")
    out = frame.copy()
    out["转债代码"] = out[code_field].map(ensure_market_suffix)
    out = out[
        out["转债代码"].str.match(
            r"^(?:110|111|113|118|123|127|128)\d{3}\.(?:SH|SZ)$",
            na=False,
        )
    ]
    if out.empty:
        raise RuntimeError(f"{report_id} 返回了数据，但 {code_field} 未解析出有效转债代码")
    return out.drop_duplicates(subset=["转债代码"]).reset_index(drop=True)


def fetch_issuing_issues() -> pd.DataFrame:
    """p04647：正在发行中的转债；f001/f002/f004/f009/f026/f042/f043。"""
    return fetch_issue_report("p04647", ISSUING_ISSUE_FIELDS, "正在发行转债")


def fetch_pending_issues() -> pd.DataFrame:
    """p04649：尚未进入发行阶段的待发行转债。"""
    return fetch_issue_report("p04649", PENDING_ISSUE_FIELDS, "待发行转债")


def ensure_market_suffix(code: Any) -> str:
    text = str(code).strip().upper()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    if text.endswith((".SH", ".SZ")):
        return text
    if text.startswith(("11", "10")):
        return f"{text}.SH"
    return f"{text}.SZ"


def fetch_details(codes: Iterable[str], as_of: date) -> pd.DataFrame:
    unique_codes = list(dict.fromkeys(ensure_market_suffix(code) for code in codes if str(code).strip()))
    if not unique_codes:
        return pd.DataFrame(columns=[*EVENT_COLUMNS, *DISPLAY_COLUMNS])

    event_result = THS_BD(
        ",".join(unique_codes),
        EVENT_FIELDS,
        ";;",
    )
    event_frame = result_frame(event_result, "发行与上市日期 THS_BD")
    if "thscode" not in event_frame.columns or len(event_frame.columns) != len(EVENT_COLUMNS) + 1:
        raise RuntimeError(f"发行与上市日期字段结构异常：{list(event_frame.columns)}")
    event_frame = event_frame.set_index("thscode")
    event_frame.columns = EVENT_COLUMNS

    # 用户指定接口：申购代码、申万行业、发行规模、评级公司、发行评级。
    display_result = THS_BD(
        ",".join(unique_codes),
        DISPLAY_FIELDS,
        f";100,{as_of.isoformat()};;;",
    )
    display_frame = result_frame(display_result, "发行日历展示字段 THS_BD")
    if "thscode" not in display_frame.columns or len(display_frame.columns) != len(DISPLAY_COLUMNS) + 1:
        raise RuntimeError(f"发行日历展示字段结构异常：{list(display_frame.columns)}")
    display_frame = display_frame.set_index("thscode")
    display_frame.columns = DISPLAY_COLUMNS

    frame = event_frame.join(display_frame, how="outer")
    frame.index = frame.index.astype(str)
    frame["发行日期"] = frame["发行日期"].map(parse_ifind_date)
    frame["上市日期"] = frame["上市日期"].map(parse_ifind_date)
    frame["发行规模"] = pd.to_numeric(frame["发行规模"], errors="coerce").map(
        lambda value: value / 100_000_000 if pd.notna(value) and abs(value) >= 1_000_000 else value
    )
    return frame


def query_trading_dates_by_offset(as_of: date, offset: int) -> list[date]:
    """以运行日为基准，调用 iFinD 返回交易日偏移序列。"""
    result = THS_Date_Offset(
        "212001",
        f"dateType:0,period:D,offset:{offset},dateFormat:0,output:sequencedate",
        as_of.isoformat(),
    )
    raw = getattr(result, "data", None)
    if isinstance(raw, str):
        values: list[Any] = [part.strip() for part in raw.split(",") if part.strip()]
    elif isinstance(raw, pd.DataFrame):
        values = raw.to_numpy().ravel().tolist()
    elif isinstance(raw, dict):
        values = []
        for item in raw.values():
            values.extend(item if isinstance(item, (list, tuple, pd.Series, pd.Index)) else [item])
    elif isinstance(raw, (list, tuple, pd.Series, pd.Index)):
        values = list(raw)
    else:
        values = []
    dates = sorted({parsed for value in values if (parsed := parse_ifind_date(value))})
    expected = abs(offset) + 1
    if len(dates) < expected:
        code = getattr(result, "errorcode", "未知")
        message = getattr(result, "errmsg", "无返回信息")
        raise RuntimeError(
            f"THS_Date_Offset({offset}) 仅返回 {len(dates)} 个交易日，"
            f"少于预期的 {expected} 个（状态码 {code}）：{message}"
        )
    return dates


def fetch_trading_dates(as_of: date, days: int) -> tuple[list[date], list[date]]:
    """返回日历展示日期及计算发行事件偏移所需的完整交易日序列。

    ``days`` 表示从运行日向后偏移的交易日数。默认偏移 5 个交易日，
    THS_Date_Offset 会返回运行日及未来 5 个交易日，共 6 个日期；表头排除
    第一个日期，只展示未来 5 个交易日。
    历史交易日同样使用负偏移获取，仅用于计算股权登记日。
    """
    future_dates = query_trading_dates_by_offset(as_of, days)
    calendar_dates = future_dates[1:]
    history_dates = query_trading_dates_by_offset(as_of, -20)
    event_trading_dates = sorted(set(history_dates + future_dates))
    return calendar_dates, event_trading_dates


def event_dates(issue_date: date | None, listed_date: date | None, trading_dates: list[date]) -> set[date]:
    events: set[date] = set()
    if listed_date:
        events.add(listed_date)
    if issue_date and issue_date in trading_dates:
        index = trading_dates.index(issue_date)
        for offset in (-1, 0, 1, 2):
            target = index + offset
            if 0 <= target < len(trading_dates):
                events.add(trading_dates[target])
    return events


def first_nonempty(primary: Any, fallback: Any = None) -> Any:
    value = clean_scalar(primary)
    if value is None or (isinstance(value, str) and not value.strip()):
        return fallback
    return value


def abbreviate_rating_agency(value: Any) -> str:
    """按固定映射缩写评级公司；未命中时取名称前四个字。"""

    text = str(first_nonempty(value, "—")).strip()
    if text == "—":
        return text
    return RATING_AGENCY_ABBREVIATIONS.get(text, text[:4])


def format_subscription_code(value: Any) -> str:
    value = clean_scalar(value)
    if value is None:
        return "—"
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text.zfill(6) if text.isdigit() else text


def build_payload(as_of: date, days: int, ths_id: str | None, ths_password: str | None) -> dict[str, Any]:
    if days <= 0:
        raise ValueError("days 必须为正整数")
    login_ifind(ths_id, ths_password)
    unlisted_codes = fetch_p05479_unlisted_codes(as_of)
    issuing = fetch_issuing_issues()
    pending = fetch_pending_issues()
    issuing_codes = issuing["转债代码"].tolist()
    pending_codes = pending["转债代码"].tolist()
    print(f"p05479 已发行未上市转债：{len(unlisted_codes)} 只")
    print(f"p05479 转债代码：{unlisted_codes or '无'}")
    print(f"p04647 正在发行转债：{len(issuing)} 只")
    print(f"p04647 转债代码：{issuing_codes or '无'}")
    print(f"p04649 待发行转债：{len(pending)} 只")
    print(f"p04649 转债代码：{pending_codes or '无'}")

    calendar_dates, trading_dates = fetch_trading_dates(as_of, days)
    issue_codes = list(dict.fromkeys([*issuing_codes, *pending_codes]))
    detail_codes = [*unlisted_codes, *issue_codes]
    details = fetch_details(detail_codes, as_of)

    bonds: list[dict[str, Any]] = []
    # 同时出现在发行中/待发行板块的代码仍按发行事件展示；只有“纯已发行未上市”
    # 的转债必须取得明确上市日期，否则不进入日历。
    unlisted_only_codes = set(unlisted_codes) - set(issue_codes)
    skipped_unlisted_without_listing: list[str] = []
    window_set = set(calendar_dates)
    for code, row in details.iterrows():
        issue_date = first_nonempty(row.get("发行日期"))
        listed_date = first_nonempty(row.get("上市日期"))
        if code in unlisted_only_codes and listed_date is None:
            skipped_unlisted_without_listing.append(code)
            continue
        if not event_dates(issue_date, listed_date, trading_dates).intersection(window_set):
            continue
        amount = first_nonempty(row.get("发行规模"))
        bonds.append(
            {
                "转债代码": code,
                "网上申购代码": format_subscription_code(row.get("网上申购代码")),
                "简称": first_nonempty(row.get("简称"), "—"),
                "所属行业": first_nonempty(row.get("所属行业"), "—"),
                "发行规模": float(amount) if amount is not None else 0.0,
                "债项评级": first_nonempty(row.get("债项评级"), "—"),
                "评级公司": abbreviate_rating_agency(row.get("评级公司")),
                "发行日期": issue_date.isoformat() if issue_date else None,
                "上市日期": listed_date.isoformat() if listed_date else None,
            }
        )

    print(
        "已发行未上市但尚无上市日期，未展示："
        f"{skipped_unlisted_without_listing or '无'}"
    )

    def sort_key(bond: dict[str, Any]) -> tuple[str, str]:
        dates = [value for value in (bond["上市日期"], bond["发行日期"]) if value]
        return (min(dates) if dates else "9999-12-31", str(bond["简称"]))

    bonds.sort(key=sort_key)
    return {
        "title": TITLE,
        "subtitle": SUBTITLE,
        "updated_date": as_of.isoformat(),
        "issue_list_sources": {
            "unlisted": "iFinD p05479",
            "issuing": "iFinD p04647",
            "pending": "iFinD p04649",
        },
        "unlisted_issue_codes": unlisted_codes,
        "issuing_issue_codes": issuing_codes,
        "pending_issue_codes": pending_codes,
        "calendar_dates": [value.isoformat() for value in calendar_dates],
        "trading_dates": [value.isoformat() for value in trading_dates],
        "colors": COLORS,
        "bonds": bonds,
    }


def find_node() -> str:
    if BUNDLED_NODE.is_file():
        return str(BUNDLED_NODE)
    node = shutil.which("node")
    if node:
        return node
    raise RuntimeError("未找到 Node.js；可通过 CODEX_BUNDLED_NODE 指定路径")


def render_payload(payload: dict[str, Any], output_dir: Path) -> dict[str, Path]:
    if not IMAGE_BUILDER.is_file():
        raise RuntimeError(f"缺少排版模块：{IMAGE_BUILDER}")
    # 兼容旧快照输入：只保留当前版日历实际使用的字段。
    allowed_bond_fields = {
        "转债代码",
        "网上申购代码",
        "简称",
        "所属行业",
        "发行规模",
        "债项评级",
        "评级公司",
        "发行日期",
        "上市日期",
    }
    normalized_bonds: list[dict[str, Any]] = []
    for bond in payload.get("bonds", []):
        normalized = {
            key: value for key, value in bond.items() if key in allowed_bond_fields
        }
        normalized["评级公司"] = abbreviate_rating_agency(normalized.get("评级公司"))
        normalized_bonds.append(normalized)
    payload["title"] = TITLE
    payload["subtitle"] = SUBTITLE
    payload["bonds"] = normalized_bonds
    as_of = datetime.strptime(payload["updated_date"], "%Y-%m-%d").date()
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = f"发行日历_{as_of.strftime('%Y%m%d')}"
    # 每次直接重建该工作簿；文件中只保留“发行日历”一个工作表。
    xlsx_path = output_dir / "可转债日历.xlsx"
    png_path = output_dir / f"{stem}.png"
    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        suffix=".json",
        prefix="bond_calendar_",
        delete=False,
    ) as temporary_json:
        json.dump(payload, temporary_json, ensure_ascii=False, indent=2)
        json_path = Path(temporary_json.name)
    try:
        subprocess.run(
            [
                find_node(),
                str(IMAGE_BUILDER),
                "--input",
                str(json_path),
                "--xlsx",
                str(xlsx_path),
                "--png",
                str(png_path),
            ],
            cwd=SCRIPT_DIR,
            check=True,
            stdout=subprocess.DEVNULL,
        )
    finally:
        json_path.unlink(missing_ok=True)
    print(f"已生成图片：{png_path}")
    print(f"Excel：{xlsx_path}")
    return {"png": png_path, "xlsx": xlsx_path}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成未来 5 个交易日的可转债发行日历图片")
    parser.add_argument("--date", default=date.today().isoformat(), help="统计日期 YYYY-MM-DD；默认今天")
    parser.add_argument(
        "--days",
        type=int,
        default=5,
        help="从运行日起向后偏移的交易日数；默认 5，日历排除运行日",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="图片和 Excel 的输出目录；默认按统计日期写入 huachuang/MMDD数据更新",
    )
    parser.add_argument("--input-json", type=Path, help="直接渲染已有 JSON，不再调用 iFinD")
    parser.add_argument("--ths-id", help="临时覆盖ifind账号.txt中的用户名")
    parser.add_argument("--ths-password", help="临时覆盖ifind账号.txt中的密码")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.input_json:
        payload = json.loads(args.input_json.resolve().read_text(encoding="utf-8"))
    else:
        as_of = datetime.strptime(args.date, "%Y-%m-%d").date()
        payload = build_payload(as_of, args.days, args.ths_id, args.ths_password)
    payload_date = datetime.strptime(payload["updated_date"], "%Y-%m-%d").date()
    output_dir = (
        args.output_dir.resolve()
        if args.output_dir is not None
        else SCRIPT_DIR / "runs" / "daily" / f"{payload_date.strftime('%m%d')}数据更新"
    )
    render_payload(payload, output_dir)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"生成失败：{exc}", file=sys.stderr)
        raise
