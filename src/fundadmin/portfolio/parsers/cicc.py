"""中金（CICC）持仓表与估值表解析器。

支持两类文件：
1. 标准持仓表（CSV / .xlsx）：包含 标的名称、合约持仓、市值（人民币）等列。
2. 估值表（.xls）：多行会计表头，包含 3199.01.01.* 收益互换科目、资产净值、单位净值。

过滤规则：
- 标准表：合约持仓 > 0。
- 估值表：提取 3199.01.01.* 叶子节点，取市值（本币）列。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from fundadmin.portfolio.parsers.common import (
    clean_ticker,
    to_float,
    to_int,
)

# 列头别名
CICC_NAME_LABELS = {"标的名称", "证券名称", "股票名称", "名称"}
CICC_CODE_LABELS = {"标的代码", "证券代码", "股票代码", "代码"}
CICC_QTY_LABELS = {"合约持仓", "持仓数量", "持仓股数", "数量"}
CICC_CNY_MV_LABELS = {"市值（人民币）", "市值(人民币)", "人民币市值", "市值"}


def _detect_header_row(df: pd.DataFrame) -> int:
    """通过关键词命中数定位表头行。"""
    target = CICC_NAME_LABELS | CICC_CODE_LABELS | CICC_QTY_LABELS | CICC_CNY_MV_LABELS
    for idx in range(min(60, len(df))):
        row_texts = {str(v or "").strip() for v in df.iloc[idx].tolist()}
        hits = sum(1 for label in target if label in row_texts)
        if hits >= 2:
            return idx
    return 0


def parse_cicc_valuation_navs(path: Path, *, sheet: str | int = 0) -> tuple[float, float]:
    """解析 CICC 估值表 .xls，提取期末单位净值和资产净值（市值口径）。

    返回:
        (unit_nav, asset_nav)

    异常:
        ValueError: 找不到对应行或数值。
    """
    df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object, engine="xlrd")
    if df.empty:
        raise ValueError(f"CICC 估值表为空: {path}")

    unit_nav: float | None = None
    asset_nav: float | None = None

    for row_idx in range(len(df)):
        first_col = str(df.iloc[row_idx, 0] or "").strip().replace(" ", "")

        # 资产净值 — 取市值本币列（索引11）
        if first_col == "资产净值":
            val = to_float(df.iloc[row_idx, 11])
            if val is not None and val > 0:
                asset_nav = val

        # 今日单位净值 / 期末单位净值
        if first_col in {"今日单位净值", "期末单位净值"}:
            val = to_float(df.iloc[row_idx, 1])
            if val is not None and val > 0:
                unit_nav = val

    # 若未找到，扫描前5行所有列中的 "单位净值:XXX" / "单位净值：XXX"
    if unit_nav is None:
        for row_idx in range(min(5, len(df))):
            for col_idx in range(min(15, len(df.columns))):
                cell = str(df.iloc[row_idx, col_idx] or "").strip().replace(" ", "")
                if "单位净值" not in cell or "累计" in cell:
                    continue
                for sep in ("：", ":"):
                    if sep in cell:
                        parts = cell.split(sep)
                        if len(parts) >= 2:
                            val = to_float(parts[-1])
                            if val is not None and val > 0:
                                unit_nav = val
                                break
                if unit_nav is not None:
                    break
            if unit_nav is not None:
                break

    if unit_nav is None:
        raise ValueError(f"在 {path} 中未找到单位净值")
    if asset_nav is None:
        raise ValueError(f"在 {path} 中未找到资产净值")

    return unit_nav, asset_nav


def parse_cicc_valuation_holdings(path: Path, *, sheet: str | int = 0) -> pd.DataFrame:
    """解析 CICC 估值表 .xls 中的收益互换持仓（3199.01.01.* 科目）。

    返回列：
        - ticker: 合约代码（如 104902--USD--EQ OTC）
        - company: 合约名称（如 中信證券-104902--USD--EQ）
        - shares: 0（swap 无股数概念）
        - market_value_cny: 市值（人民币，取本币列）
        - source_file: 来源文件名
    """
    df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object, engine="xlrd")
    if df.empty:
        raise ValueError(f"CICC 估值表为空: {path}")

    rows: list[dict[str, Any]] = []
    for row_idx in range(len(df)):
        code = str(df.iloc[row_idx, 0] or "").strip()
        # 只取 3199.01.01.* 叶子节点（排除汇总行 3199.01.01）
        if not code.startswith("3199.01.01."):
            continue

        name = str(df.iloc[row_idx, 1] or "").strip()
        mv = to_float(df.iloc[row_idx, 11])  # 市值本币列

        if mv is None:
            continue

        rows.append(
            {
                "ticker": code,
                "company": name,
                "shares": 0,
                "market_value_cny": mv,
                "source_file": path.name,
            }
        )

    if not rows:
        return pd.DataFrame(columns=["ticker", "company", "shares", "market_value_cny", "source_file"])

    return pd.DataFrame(rows)


def parse_cicc_holdings(path: Path, *, sheet: str | int = 0) -> pd.DataFrame:
    """解析中金持仓表，返回标准化 DataFrame。

    返回列：
        - ticker: 统一代码（如 AMZN）
        - company: 公司名称
        - shares: 持仓股数
        - market_value_cny: 市值（人民币）
        - source_file: 来源文件名
    """
    suffix = path.suffix.lower()
    if suffix == ".csv":
        df_raw = pd.read_csv(path, header=None, dtype=object, encoding="utf-8-sig")
    elif suffix == ".xls":
        df_raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object, engine="xlrd")
    else:
        df_raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object, engine="openpyxl")

    if df_raw.empty:
        raise ValueError(f"中金持仓表为空: {path}")

    header_row = _detect_header_row(df_raw)
    headers = [str(v or "").strip() for v in df_raw.iloc[header_row].tolist()]
    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")

    # 映射列名
    col_map: dict[str, str] = {}
    for col in df.columns:
        text = str(col).strip()
        if text in CICC_NAME_LABELS:
            col_map[col] = "company"
        elif text in CICC_CODE_LABELS:
            col_map[col] = "code"
        elif text in CICC_QTY_LABELS:
            col_map[col] = "shares"
        elif text in CICC_CNY_MV_LABELS:
            col_map[col] = "market_value_cny"

    df = df.rename(columns=col_map)

    if "company" not in df.columns:
        raise ValueError(f"中金持仓表缺少公司名称列: {path}")
    if "shares" not in df.columns:
        raise ValueError(f"中金持仓表缺少持仓数量列: {path}")

    # 过滤空仓位
    df["shares_num"] = df["shares"].map(to_float).fillna(0.0)
    df = df[df["shares_num"] > 0].copy()
    if df.empty:
        return pd.DataFrame(columns=["ticker", "company", "shares", "market_value_cny", "source_file"])

    out = pd.DataFrame()
    out["company"] = df["company"].astype(str).str.strip()
    out["ticker"] = df["code"].map(clean_ticker) if "code" in df.columns else ""
    out["shares"] = df["shares"].map(to_int)
    out["market_value_cny"] = df["market_value_cny"].map(to_float) if "market_value_cny" in df.columns else None
    out["source_file"] = path.name

    return out


# CICC 估值报告 .xlsx（如 2026-04-16弘运盛泰铂金2号私募证券投资基金.xlsx）持仓 sheet 列名
CICC_VAL_XLSX_CODE_COL = "标的代码"
CICC_VAL_XLSX_NAME_COL = "标的名称"
CICC_VAL_XLSX_QTY_COL = "合约持仓"
CICC_VAL_XLSX_MV_CNY_COL = "市值（人民币）"


def parse_cicc_valuation_xlsx_holdings(path: Path, *, sheet: str = "持仓") -> pd.DataFrame:
    """解析 CICC 估值报告 .xlsx 中的持仓 sheet。

    文件特征：文件名形如 {date}弘运盛泰{product}私募证券投资基金.xlsx，
    内含 "持仓" sheet，列包含 标的代码、标的名称、合约持仓、市值（人民币）等。

    返回列：
        - ticker: 标的代码（如 SATS.US）
        - company: 标的名称
        - shares: 合约持仓股数
        - market_value_cny: 市值（人民币）
        - source_file: 来源文件名
    """
    df = pd.read_excel(path, sheet_name=sheet, dtype=object, engine="openpyxl")
    if df.empty:
        return pd.DataFrame(columns=["ticker", "company", "shares", "market_value_cny", "source_file"])

    # 标准化列名
    col_map: dict[str, str] = {}
    for col in df.columns:
        text = str(col).strip()
        if text == CICC_VAL_XLSX_CODE_COL:
            col_map[col] = "ticker"
        elif text == CICC_VAL_XLSX_NAME_COL:
            col_map[col] = "company"
        elif text == CICC_VAL_XLSX_QTY_COL:
            col_map[col] = "shares"
        elif text == CICC_VAL_XLSX_MV_CNY_COL:
            col_map[col] = "market_value_cny"

    # 容错：如果列名没匹配到，尝试通过位置映射（前 10 列内）
    if "ticker" not in col_map:
        for idx, col in enumerate(df.columns):
            if idx == 0:
                col_map[col] = "ticker"
            elif idx == 1:
                col_map[col] = "company"
            elif idx == 3:
                col_map[col] = "shares"
            elif idx == 9:
                col_map[col] = "market_value_cny"

    df = df.rename(columns=col_map)

    if "ticker" not in df.columns:
        raise ValueError(f"CICC 估值报告持仓 sheet 缺少标的代码列: {path}")
    if "company" not in df.columns:
        raise ValueError(f"CICC 估值报告持仓 sheet 缺少标的名称列: {path}")
    if "shares" not in df.columns:
        raise ValueError(f"CICC 估值报告持仓 sheet 缺少合约持仓列: {path}")

    df = df.dropna(how="all")
    df["shares_num"] = df["shares"].map(to_float).fillna(0.0)
    df = df[df["shares_num"] > 0].copy()
    if df.empty:
        return pd.DataFrame(columns=["ticker", "company", "shares", "market_value_cny", "source_file"])

    out = pd.DataFrame()
    out["ticker"] = df["ticker"].astype(str).str.strip()
    out["company"] = df["company"].astype(str).str.strip()
    out["shares"] = df["shares"].map(to_int)
    out["market_value_cny"] = df["market_value_cny"].map(to_float) if "market_value_cny" in df.columns else None
    out["source_file"] = path.name
    return out
