# -*- coding: utf-8 -*-

"""匹配正股前十大股东与转债前十大持有人并输出指定期限 Excel。"""

from __future__ import annotations

import argparse
from configparser import ConfigParser
from dataclasses import dataclass
from datetime import datetime
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import Any, Iterable

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
IFIND_CREDENTIAL_FILE = ROOT / "private" / "ifind账号.txt"
WORKBOOK_BUILDER = ROOT / "scripts" / "build_major_shareholder_cb_holdings_workbook.mjs"
THS_LOGIN_OK_CODES = (0, -201)
MATURITY_MIN_YEARS = 5.0
MATURITY_MAX_YEARS = 5.5
HOLDER_RANKS = range(1, 11)

CB_BASIC_FIELDS = (
    "ths_convertible_debt_short_name_cbond;"
    "ths_stock_code_cbond;"
    "ths_stock_short_name_cbond;"
    "ths_issue_method_cbond;"
    "ths_trading_status_bond"
)
CB_BASIC_COLUMNS = ["转债简称", "正股代码", "正股简称", "发行方式", "交易状态"]


@dataclass(frozen=True)
class HolderMatch:
    rank: int | None
    ratio: float | None
    status: str


def load_ifind_credentials(path: Path = IFIND_CREDENTIAL_FILE) -> tuple[str, str]:
    """从项目 private/ifind账号.txt 读取公用 iFinD 账号。"""
    if not path.is_file():
        raise FileNotFoundError(f"未找到 iFinD 账号文件：{path}")

    config = ConfigParser(interpolation=None)
    config.read(path, encoding="utf-8")
    username = config.get("ifind", "username", fallback="").strip()
    password = config.get("ifind", "password", fallback="").strip()
    if not username or not password:
        raise RuntimeError(f"{path} 中的 [ifind] username 或 password 为空")
    return username, password


def _as_dataframe(result: Any, label: str) -> pd.DataFrame:
    data = getattr(result, "data", None)
    if data is None:
        error_code = getattr(result, "errorcode", "未知")
        error_message = getattr(result, "errmsg", "未返回数据")
        raise RuntimeError(f"{label}失败（{error_code}）：{error_message}")
    if not isinstance(data, pd.DataFrame):
        data = pd.DataFrame(data)
    if data.empty:
        raise RuntimeError(f"{label}未返回数据")
    return data.copy()


def _clean_code_list(codes: Iterable[str]) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for code in codes:
        value = str(code).strip()
        if value and value not in seen:
            cleaned.append(value)
            seen.add(value)
    return cleaned


def _clean_name(value: Any) -> str | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    return text or None


def _clean_number(value: Any) -> float | None:
    number = pd.to_numeric(value, errors="coerce")
    return None if pd.isna(number) else float(number)


def match_major_shareholder(
    major_holder_name: Any,
    top10_names: Iterable[Any],
    top10_ratios: Iterable[Any],
) -> HolderMatch:
    """按名称去除首尾空格后精确匹配，避免模糊匹配误认同名主体。"""
    target = _clean_name(major_holder_name)
    names = [_clean_name(value) for value in top10_names]
    ratios = list(top10_ratios)

    if target is None:
        return HolderMatch(None, None, "大股东名称缺失")
    if not any(name is not None for name in names):
        return HolderMatch(None, None, "前十大持有人数据缺失")

    for rank, name in enumerate(names, start=1):
        if name == target:
            ratio = _clean_number(ratios[rank - 1]) if rank <= len(ratios) else None
            status = "匹配成功" if ratio is not None else "名称匹配但持债比例缺失"
            return HolderMatch(rank, ratio, status)
    return HolderMatch(None, None, "前十大中未找到")


def format_holding_status(match: HolderMatch) -> str:
    if match.status == "匹配成功":
        return f"转债第{match.rank}名，持债{match.ratio:.2f}%"
    if match.status == "名称匹配但持债比例缺失":
        return f"转债第{match.rank}名，持债比例缺失"
    if match.status == "前十大中未找到":
        return "未进入转债前十大"
    return match.status


def login_ifind() -> int:
    from iFinDPy import THS_GetErrorInfo, THS_iFinDLogin

    username, password = load_ifind_credentials()
    code = THS_iFinDLogin(username, password)
    if code not in THS_LOGIN_OK_CODES:
        try:
            detail = THS_GetErrorInfo(code)
        except Exception:
            detail = ""
        raise RuntimeError(f"iFinD 登录失败（状态码 {code}）：{detail}")
    print("iFinD 登录成功" if code == 0 else "iFinD 账号已登录，复用现有会话")
    return code


def get_last_trade_date(as_of: str | None = None) -> str:
    from iFinDPy import THS_Date_Offset

    query_date = as_of or datetime.now().strftime("%Y-%m-%d")
    result = THS_Date_Offset(
        "212001",
        "dateType:0,period:D,offset:0,dateFormat:0,output:singledate",
        query_date,
    )
    value = getattr(result, "data", None)
    if value is None:
        raise RuntimeError("THS_Date_Offset 未返回交易日，请确认 iFinD 登录状态")
    return pd.Timestamp(value).strftime("%Y-%m-%d")


def fetch_current_cb_universe(last_date: str) -> pd.DataFrame:
    """获取指定交易日未到期、非定向、非 NQ、非终止上市的转债。"""
    from iFinDPy import THS_BD, THS_DR

    edate = pd.Timestamp(last_date).strftime("%Y%m%d")
    raw_list = _as_dataframe(
        THS_DR(
            "p00570",
            f"jyzt=未到期;sfdb=全部;jysc=全部;edate={edate}",
            "jydm:Y,jydm_mc:Y,p00570_f001:Y,p00570_f019:Y",
            "format:dataframe",
        ),
        f"获取 {last_date} 可转债列表",
    )
    if "jydm" not in raw_list.columns:
        raise RuntimeError("可转债列表缺少 jydm 字段")
    codes = _clean_code_list(raw_list["jydm"].astype(str))
    if not codes:
        raise RuntimeError(f"{last_date} 没有可用的转债代码")

    basic = _as_dataframe(
        THS_BD(",".join(codes), CB_BASIC_FIELDS, ";;;;"),
        "获取可转债基础信息",
    )
    if "thscode" not in basic.columns or len(basic.columns) - 1 != len(CB_BASIC_COLUMNS):
        raise RuntimeError(f"可转债基础信息字段数量异常：{list(basic.columns)}")
    basic = basic.set_index("thscode")
    basic.columns = CB_BASIC_COLUMNS
    basic.index = basic.index.astype(str)
    basic.index.name = "转债代码"
    basic = basic[~basic["发行方式"].astype(str).str.contains("定向", na=False)]
    basic = basic[~basic.index.str.contains("NQ", na=False)]
    basic = basic[~basic["交易状态"].astype(str).str.contains("终止上市", na=False)]
    return basic.sort_index()


def fetch_requested_cb_universe(codes: list[str]) -> pd.DataFrame:
    from iFinDPy import THS_BD

    basic = _as_dataframe(
        THS_BD(",".join(codes), CB_BASIC_FIELDS, ";;;;"),
        "获取指定可转债基础信息",
    )
    if "thscode" not in basic.columns or len(basic.columns) - 1 != len(CB_BASIC_COLUMNS):
        raise RuntimeError(f"可转债基础信息字段数量异常：{list(basic.columns)}")
    basic = basic.set_index("thscode")
    basic.columns = CB_BASIC_COLUMNS
    basic.index = basic.index.astype(str)
    basic.index.name = "转债代码"
    return basic.reindex(codes)


def fetch_remaining_maturity(codes: list[str], last_date: str) -> pd.Series:
    """获取基准日剩余期限，单位为年。"""
    from iFinDPy import THS_BD

    code_text = ",".join(codes)
    raw = _as_dataframe(
        THS_BD(code_text, "ths_remain_duration_y_bond", last_date),
        "获取剩余期限",
    ).set_index("thscode")
    maturity = pd.to_numeric(raw.iloc[:, 0], errors="coerce").reindex(codes)
    maturity.name = "剩余期限（年）"
    maturity.index.name = "转债代码"
    return maturity


def filter_maturity_range(
    basic: pd.DataFrame,
    maturity: pd.Series,
    minimum: float = MATURITY_MIN_YEARS,
    maximum: float = MATURITY_MAX_YEARS,
) -> pd.DataFrame:
    out = basic.copy()
    out["剩余期限（年）"] = pd.to_numeric(maturity.reindex(out.index), errors="coerce")
    return out[out["剩余期限（年）"].between(minimum, maximum, inclusive="both")].copy()


def fetch_holder_data(
    codes: list[str],
    last_date: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """返回正股前十大股东及转债前十大持有人的名称和比例。"""
    from iFinDPy import THS_BD

    code_text = ",".join(codes)
    stock_holder_names: dict[str, pd.Series] = {}
    stock_holder_ratios: dict[str, pd.Series] = {}
    for rank in HOLDER_RANKS:
        stock_df = _as_dataframe(
            THS_BD(
                code_text,
                "ths_major_shareholder_name_bond;ths_big_holder_held_ratio_bond",
                f"{last_date},{rank};{last_date},{rank}",
            ),
            f"获取正股第 {rank} 名大股东信息",
        ).set_index("thscode")
        if len(stock_df.columns) != 2:
            raise RuntimeError(f"正股第 {rank} 名大股东字段数量异常：{list(stock_df.columns)}")
        stock_holder_names[f"NO{rank}"] = stock_df.iloc[:, 0].reindex(codes)
        stock_holder_ratios[f"NO{rank}"] = stock_df.iloc[:, 1].reindex(codes)

    holder_names: dict[str, pd.Series] = {}
    holder_ratios: dict[str, pd.Series] = {}
    for rank in HOLDER_RANKS:
        name_df = _as_dataframe(
            THS_BD(code_text, "ths_holder_name_cbond", f"{last_date},{rank}"),
            f"获取第 {rank} 名转债持有人名称",
        ).set_index("thscode")
        ratio_df = _as_dataframe(
            THS_BD(code_text, "ths_holder_held_ratio_cbond", f"{last_date},{rank}"),
            f"获取第 {rank} 名转债持有人持债比例",
        ).set_index("thscode")
        holder_names[f"第{rank}名"] = name_df.iloc[:, 0].reindex(codes)
        holder_ratios[f"第{rank}名"] = ratio_df.iloc[:, 0].reindex(codes)

    return (
        pd.DataFrame(stock_holder_names),
        pd.DataFrame(stock_holder_ratios),
        pd.DataFrame(holder_names),
        pd.DataFrame(holder_ratios),
    )


def build_output_records(
    basic: pd.DataFrame,
    stock_holder_names: pd.DataFrame,
    stock_holder_ratios: pd.DataFrame,
    bond_holder_names: pd.DataFrame,
    bond_holder_ratios: pd.DataFrame,
    last_date: str,
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []

    for code in basic.index.astype(str):
        bond_names = bond_holder_names.loc[code].tolist()
        bond_ratios = bond_holder_ratios.loc[code].tolist()
        record: dict[str, Any] = {
            "查询日期": last_date,
            "转债代码": code,
            "转债简称": _clean_name(basic.at[code, "转债简称"]),
            "正股代码": _clean_name(basic.at[code, "正股代码"]),
            "正股简称": _clean_name(basic.at[code, "正股简称"]),
            "剩余期限（年）": _clean_number(basic.at[code, "剩余期限（年）"]),
        }
        for rank in HOLDER_RANKS:
            stock_name = stock_holder_names.at[code, f"NO{rank}"]
            match = match_major_shareholder(stock_name, bond_names, bond_ratios)
            record[f"NO{rank}大股东"] = _clean_name(stock_name)
            record[f"NO{rank}持股比例（%）"] = _clean_number(
                stock_holder_ratios.at[code, f"NO{rank}"]
            )
            record[f"NO{rank}持债情况"] = format_holding_status(match)
        records.append(record)

    return records


def _artifact_runtime_paths() -> tuple[Path, Path]:
    """定位 Codex 随附的 Node 和 artifact-tool；可用环境变量覆盖。"""
    node_override = os.environ.get("CODEX_ARTIFACT_NODE", "").strip()
    modules_override = os.environ.get("CODEX_ARTIFACT_NODE_MODULES", "").strip()
    dependency_root = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
    )
    node_path = Path(node_override) if node_override else dependency_root / "node" / "bin" / "node.exe"
    modules_path = Path(modules_override) if modules_override else dependency_root / "node" / "node_modules"
    if not node_path.is_file():
        fallback = shutil.which("node")
        if fallback:
            node_path = Path(fallback)
    if not node_path.is_file():
        raise FileNotFoundError("未找到 Node.js；可通过 CODEX_ARTIFACT_NODE 指定路径")
    if not (modules_path / "@oai" / "artifact-tool").exists():
        raise FileNotFoundError(
            "未找到 @oai/artifact-tool；可通过 CODEX_ARTIFACT_NODE_MODULES 指定 node_modules 路径"
        )
    return node_path, modules_path


def _create_directory_link(link: Path, target: Path) -> None:
    try:
        link.symlink_to(target, target_is_directory=True)
        return
    except OSError:
        if os.name != "nt":
            raise
    result = subprocess.run(
        ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"创建 node_modules 目录链接失败：{result.stderr or result.stdout}")


def export_workbook(payload: dict[str, Any], output_path: Path, qa_dir: Path | None = None) -> None:
    if not WORKBOOK_BUILDER.is_file():
        raise FileNotFoundError(f"未找到 Excel 构建器：{WORKBOOK_BUILDER}")
    node_path, modules_path = _artifact_runtime_paths()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="cb_holder_workbook_") as temp_name:
        temp_dir = Path(temp_name)
        payload_path = temp_dir / "holder_data.json"
        builder_path = temp_dir / WORKBOOK_BUILDER.name
        payload_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False),
            encoding="utf-8",
        )
        shutil.copy2(WORKBOOK_BUILDER, builder_path)
        _create_directory_link(temp_dir / "node_modules", modules_path)

        command = [
            str(node_path),
            str(builder_path),
            "--input",
            str(payload_path),
            "--output",
            str(output_path),
        ]
        if qa_dir is not None:
            qa_dir.mkdir(parents=True, exist_ok=True)
            command.extend(["--qa-dir", str(qa_dir)])
        subprocess.run(command, cwd=temp_dir, check=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="获取5.0—5.5年转债的正股前十大股东持债情况")
    parser.add_argument("--date", help="查询基准日，格式 YYYY-MM-DD；默认取今天对应的最近交易日")
    parser.add_argument("--codes", help="可选，逗号分隔的转债代码；默认查询全部未到期转债")
    parser.add_argument("--output", type=Path, help="输出 Excel 路径")
    parser.add_argument("--qa-dir", type=Path, help="可选，保存 Excel 渲染校验图片的目录")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    login_code: int | None = None
    try:
        login_code = login_ifind()
        last_date = get_last_trade_date(args.date)
        print(f"查询交易日：{last_date}")

        requested_codes = _clean_code_list((args.codes or "").replace("，", ",").split(","))
        basic = (
            fetch_requested_cb_universe(requested_codes)
            if requested_codes
            else fetch_current_cb_universe(last_date)
        )
        all_codes = basic.index.astype(str).tolist()
        maturity = fetch_remaining_maturity(all_codes, last_date)
        basic = filter_maturity_range(basic, maturity)
        codes = basic.index.astype(str).tolist()
        print(f"剩余期限5.0—5.5年转债数量：{len(codes)}")

        if codes:
            stock_names, stock_ratios, bond_names, bond_ratios = fetch_holder_data(codes, last_date)
            records = build_output_records(
                basic,
                stock_names,
                stock_ratios,
                bond_names,
                bond_ratios,
                last_date,
            )
        else:
            records = []

        matched_cells = sum(
            str(record[f"NO{rank}持债情况"]).startswith("转债第")
            for record in records
            for rank in HOLDER_RANKS
        )
        output_path = args.output or (
            ROOT
            / "outputs"
            / "大股东持债情况"
            / f"【华创固收】大股东持债情况_剩余期限5-5.5年-{last_date.replace('-', '')}.xlsx"
        )
        payload = {
            "metadata": {
                "query_date": last_date,
                "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "source": "同花顺 iFinD",
                "universe": "剩余期限5.0—5.5年（含边界）的转债",
                "matching_rule": "正股前十大股东名称去除首尾空格后，与转债前十大持有人名称精确匹配",
                "total_bonds": len(records),
                "matched_shareholder_slots": matched_cells,
            },
            "records": records,
        }
        export_workbook(payload, output_path.resolve(), args.qa_dir.resolve() if args.qa_dir else None)
        print(f"Excel 已输出：{output_path.resolve()}")
        print(f"NO1—NO10 共匹配到 {matched_cells} 个大股东持债记录")
        return 0
    finally:
        if login_code == 0:
            try:
                from iFinDPy import THS_iFinDLogout

                THS_iFinDLogout()
            except Exception as exc:
                print(f"[警告] iFinD 退出登录失败：{exc}")


if __name__ == "__main__":
    raise SystemExit(main())
