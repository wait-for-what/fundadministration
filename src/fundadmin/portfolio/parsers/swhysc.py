"""申万宏源（swhysc）资产估值表解析器。

结构特征：
- Sheet 名通常为"估值表明细"。
- 列头：科目代码、科目名称、数量、单位成本、成本、成本占净值%、市价、市值、市值占净值%、估值增值
- NAV 行：包含"基金资产净值"，取"市值"列。
- 个股行：科目代码为 14 位数字（如 11020101600176），数量 > 0，市值 > 0。
  - 后 6 位为股票代码。
  - 第 5-6 位标识交易所：01=上交所.SH，31=深交所.SZ，41=创业板.SZ，C1=科创板.SH，03=ETF。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from fundadmin.portfolio.parsers.common import to_float, to_int

# NAV 行标签（支持带冒号变体）
NAV_ROW_LABELS = {"基金资产净值", "基金资产净值:", "资产净值", "资产净值:", "净资产", "净资产:", "今日资产净值"}
# 单位净值行标签
UNIT_NAV_LABELS = {"单位净值", "单位净值:", "期末单位净值", "期末单位净值:"}
# 个股科目代码前缀（股票投资和基金投资）
STOCK_PREFIXES = {"1102", "1105"}
# 减值准备标识（第7-8位为99）
IMPairment_MARKER = "99"


def _detect_header_row(df: pd.DataFrame) -> int:
    for idx in range(min(10, len(df))):
        row_texts = {str(v or "").strip() for v in df.iloc[idx].tolist()}
        if "科目代码" in row_texts and "科目名称" in row_texts:
            return idx
    return 3


def _extract_stock_code(full_code: str) -> tuple[str, str]:
    """从 14 位科目代码中提取股票代码和交易所后缀。

    返回 (ticker, exchange_suffix)
    """
    if len(full_code) < 10:
        return full_code, ""

    # 后6位是股票代码
    stock_code = full_code[-6:]

    # 第5-6位（索引4:6）标识子市场
    segment = full_code[4:6] if len(full_code) >= 6 else ""
    if segment == "01" or segment.upper() == "C1":
        return stock_code, ".SH"
    elif segment in {"31", "41", "03"}:
        return stock_code, ".SZ"
    else:
        return stock_code, ""


def parse_swhysc_navs(path: Path) -> tuple[float | None, float | None]:
    """从申万宏源估值表中提取单位净值和资产净值（均优先取市值口径）。

    返回:
        (unit_nav, asset_nav) — 任一值可能为 None，表示未找到。
    """
    df = pd.read_excel(path, sheet_name=0, header=None, dtype=object, engine="openpyxl")
    if df.empty:
        return None, None

    unit_nav: float | None = None
    asset_nav: float | None = None

    for row_idx in range(len(df)):
        first_col = str(df.iloc[row_idx, 0] or "").strip().replace(" ", "")

        # 资产净值 — 优先取市值列（索引7），其次成本列（索引4）
        if first_col in NAV_ROW_LABELS:
            for col_idx in (7, 4, 8):
                val = to_float(df.iloc[row_idx, col_idx])
                if val is not None and val > 0:
                    asset_nav = val
                    break

        # 期末单位净值 — 优先取索引7，其次索引1
        if "期末单位净值" in first_col:
            for col_idx in (7, 1, 4):
                val = to_float(df.iloc[row_idx, col_idx])
                if val is not None and val > 0:
                    unit_nav = val
                    break

    # 若未找到期末单位净值行，扫描前5行中的 "单位净值" 单元格（如 Row 2 的 "单位净值：1.194"）
    if unit_nav is None:
        for row_idx in range(min(5, len(df))):
            for col_idx in range(min(12, len(df.columns))):
                cell = str(df.iloc[row_idx, col_idx] or "").strip().replace(" ", "")
                if "单位净值" not in cell or "累计" in cell:
                    continue
                # 尝试从同单元格提取 "单位净值：1.194"
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

    return unit_nav, asset_nav


def parse_swhysc_nav(path: Path) -> float:
    """从申万宏源估值表中提取资产净值（市值口径），保持向后兼容。

    异常:
        ValueError: 找不到 NAV 行或数值。
    """
    _, asset_nav = parse_swhysc_navs(path)
    if asset_nav is not None:
        return asset_nav
    raise ValueError(f"在 {path} 中未找到资产净值")


def parse_swhysc_holdings(path: Path) -> pd.DataFrame:
    """解析申万宏源估值表中的国内持仓明细。

    返回列：
        - ticker: 统一代码（如 600176.SH）
        - company: 公司名称
        - shares: 持仓股数
        - market_value_cny: 市值（人民币，元）
        - source_file: 来源文件名
    """
    df_raw = pd.read_excel(path, sheet_name=0, header=None, dtype=object, engine="openpyxl")
    if df_raw.empty:
        raise ValueError(f"估值表为空: {path}")

    header_row = _detect_header_row(df_raw)
    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = df_raw.iloc[header_row]

    # 重命名列便于访问
    col_map: dict[str, str] = {}
    for col in df.columns:
        text = str(col).strip()
        if text == "科目代码":
            col_map[col] = "code"
        elif text == "科目名称":
            col_map[col] = "name"
        elif text == "数量":
            col_map[col] = "shares"
        elif text == "市值":
            col_map[col] = "market_value"

    df = df.rename(columns=col_map)

    if "code" not in df.columns or "name" not in df.columns:
        raise ValueError(f"申万估值表缺少必要列: {path}")

    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        code = str(row.get("code", "")).strip()
        if not code.isdigit() or len(code) < 10:
            continue
        # 排除减值准备行（第7-8位为99）
        if len(code) >= 8 and code[6:8] == IMPairment_MARKER:
            continue
        # 排除非股票/基金科目
        if not any(code.startswith(prefix) for prefix in STOCK_PREFIXES):
            continue

        shares = to_int(row.get("shares"))
        mv = to_float(row.get("market_value"))
        if shares is None or shares <= 0 or mv is None or mv <= 0:
            continue

        ticker, suffix = _extract_stock_code(code)
        if suffix:
            ticker = ticker + suffix

        rows.append(
            {
                "ticker": ticker,
                "company": str(row.get("name", "")).strip(),
                "shares": shares,
                "market_value_cny": mv,
                "source_file": path.name,
            }
        )

    if not rows:
        return pd.DataFrame(columns=["ticker", "company", "shares", "market_value_cny", "source_file"])

    return pd.DataFrame(rows)
