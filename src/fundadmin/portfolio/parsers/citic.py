"""中信（CITIC）底层资产表 + Balance 汇率解析器。

支持 USD 和 HKD 账户：
- Underlying.csv: 提取持仓明细。
- Balance.csv: 提取汇率（交易货币/结算货币）。

市值换算：
    市值（人民币） = 标的市值（交易货币） × 汇率
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

# Underlying 列头别名
CITIC_NAME_LABELS = {"证券名称", "标的名称", "名称"}
CITIC_CODE_LABELS = {"证券代码", "标的代码", "代码"}
CITIC_QTY_LABELS = {"标的名义数量加总", "名义数量加总", "持仓数量", "数量"}
CITIC_MV_LABELS = {"标的市值（交易货币）", "市值（交易货币）", "标的市值", "市值"}
# 成本相关列头别名（交易货币口径）
CITIC_UNIT_COST_LABELS = {"平均单位持仓成本（交易货币）", "平均单位持仓成本(交易货币)", "平均单位持仓成本"}
CITIC_COST_VALUE_LABELS = {
    "该标的持仓成本加总的近似值（交易货币）",
    "该标的持仓成本加总的近似值(交易货币)",
    "该标的持仓成本加总",
}
CITIC_CCY_LABELS = {"交易货币", "计价货币", "币种"}

# Balance 汇率列头别名
CITIC_FX_LABELS = {"汇率（交易货币/结算货币）", "汇率", "汇率(交易货币/结算货币)"}


def _detect_header_row(df: pd.DataFrame, target_labels: set[str], min_hits: int = 2) -> int:
    for idx in range(min(60, len(df))):
        row_texts = {str(v or "").strip() for v in df.iloc[idx].tolist()}
        hits = sum(1 for label in target_labels if label in row_texts)
        if hits >= min_hits:
            return idx
    return 0


def parse_citic_fx(balance_path: Path, *, sheet: str | int = 0) -> float:
    """从 Balance 表（CSV 或 Excel）中读取汇率。

    参数:
        balance_path: Balance 文件路径。
        sheet: 对于 Excel 文件，指定 sheet 名或索引。

    返回:
        汇率数值（如 6.8187）。

    异常:
        ValueError: 找不到汇率列或数值。
    """
    suffix = balance_path.suffix.lower()
    if suffix == ".csv":
        df_raw = pd.read_csv(balance_path, header=None, dtype=object, encoding="utf-8-sig")
    elif suffix in {".xlsx", ".xlsm"}:
        df_raw = pd.read_excel(balance_path, sheet_name=sheet, header=None, dtype=object, engine="openpyxl")
    elif suffix == ".xls":
        df_raw = pd.read_excel(balance_path, sheet_name=sheet, header=None, dtype=object, engine="xlrd")
    else:
        df_raw = pd.read_excel(balance_path, sheet_name=sheet, header=None, dtype=object)
    if df_raw.empty:
        raise ValueError(f"Balance 表为空: {balance_path}")

    header_row = _detect_header_row(df_raw, CITIC_FX_LABELS, min_hits=1)
    headers = [str(v or "").strip() for v in df_raw.iloc[header_row].tolist()]
    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")

    # 找到汇率列
    fx_col = None
    for col in df.columns:
        if col in CITIC_FX_LABELS:
            fx_col = col
            break

    if fx_col is None:
        # 兜底：找包含"汇率"字样的列
        for col in df.columns:
            if "汇率" in str(col):
                fx_col = col
                break

    if fx_col is None:
        raise ValueError(f"在 {balance_path} 中未找到汇率列")

    # 取第一个非空数值
    for val in df[fx_col]:
        num = to_float(val)
        if num is not None and num > 0:
            return num

    raise ValueError(f"在 {balance_path} 中未读取到有效汇率数值")


def parse_citic_underlying(underlying_path: Path, fx: float, *, sheet: str | int = 0) -> pd.DataFrame:
    """解析中信底层资产表，返回标准化 DataFrame。

    参数:
        underlying_path: Underlying 文件路径（CSV 或 Excel）。
        fx: 汇率（交易货币/结算货币），由 parse_citic_fx 提供。
        sheet: 对于 Excel 文件，指定 sheet 名或索引。

    返回列：
        - ticker: 统一代码（如 TSLA、0286）
        - company: 公司名称
        - shares: 持仓股数
        - market_value_cny: 市值（人民币）
        - cost_price_local: 平均单位持仓成本（交易货币，每股）
        - cost_value_local: 该标的持仓成本加总（交易货币）
        - cost_ccy: 交易货币（USD/HKD）
        - source_file: 来源文件名

    说明：
        - 成本仅对真实持仓（标的市值 > 0）赋值；方向为 0 的占位行（市值为 0、
          平均成本为残值）不赋成本，避免污染成本看板。
    """
    suffix = underlying_path.suffix.lower()
    if suffix == ".csv":
        df_raw = pd.read_csv(underlying_path, header=None, dtype=object, encoding="utf-8-sig")
    elif suffix in {".xlsx", ".xlsm"}:
        df_raw = pd.read_excel(underlying_path, sheet_name=sheet, header=None, dtype=object, engine="openpyxl")
    elif suffix == ".xls":
        df_raw = pd.read_excel(underlying_path, sheet_name=sheet, header=None, dtype=object, engine="xlrd")
    else:
        df_raw = pd.read_excel(underlying_path, sheet_name=sheet, header=None, dtype=object)

    if df_raw.empty:
        raise ValueError(f"中信底层资产表为空: {underlying_path}")

    target = CITIC_NAME_LABELS | CITIC_CODE_LABELS | CITIC_QTY_LABELS | CITIC_MV_LABELS
    header_row = _detect_header_row(df_raw, target)
    headers = [str(v or "").strip() for v in df_raw.iloc[header_row].tolist()]
    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")

    col_map: dict[str, str] = {}
    for col in df.columns:
        text = str(col).strip()
        if text in CITIC_NAME_LABELS:
            col_map[col] = "company"
        elif text in CITIC_CODE_LABELS:
            col_map[col] = "code"
        elif text in CITIC_QTY_LABELS:
            col_map[col] = "shares"
        elif text in CITIC_MV_LABELS:
            col_map[col] = "market_value"
        elif text in CITIC_UNIT_COST_LABELS:
            col_map[col] = "unit_cost"
        elif text in CITIC_COST_VALUE_LABELS:
            col_map[col] = "cost_value"
        elif text in CITIC_CCY_LABELS and "ccy" not in col_map.values():
            col_map[col] = "ccy"

    df = df.rename(columns=col_map)

    if "company" not in df.columns:
        raise ValueError(f"中信底层资产表缺少公司名称列: {underlying_path}")
    if "shares" not in df.columns:
        raise ValueError(f"中信底层资产表缺少持仓数量列: {underlying_path}")

    out_cols = [
        "ticker", "company", "shares", "market_value_cny",
        "cost_price_local", "cost_value_local", "cost_ccy", "source_file",
    ]

    # 过滤空仓位
    df["shares_num"] = df["shares"].map(to_float).fillna(0.0)
    df = df[df["shares_num"] > 0].copy()
    if df.empty:
        return pd.DataFrame(columns=out_cols)

    mv_local = df["market_value"].map(to_float) if "market_value" in df.columns else pd.Series([None] * len(df))

    out = pd.DataFrame()
    out["company"] = df["company"].astype(str).str.strip()
    out["ticker"] = df["code"].map(clean_ticker) if "code" in df.columns else ""
    out["shares"] = df["shares"].map(to_int)
    out["market_value_cny"] = mv_local * fx

    # 成本（交易货币口径），仅对真实持仓（标的市值 > 0）赋值。
    real = mv_local.fillna(0.0) > 0
    unit_cost = df["unit_cost"].map(to_float) if "unit_cost" in df.columns else pd.Series([None] * len(df))
    cost_value = df["cost_value"].map(to_float) if "cost_value" in df.columns else pd.Series([None] * len(df))
    out["cost_price_local"] = unit_cost.where(real, other=None)
    out["cost_value_local"] = cost_value.where(real, other=None)
    if "ccy" in df.columns:
        out["cost_ccy"] = df["ccy"].astype(str).str.strip().where(real, other=None)
    else:
        out["cost_ccy"] = None
    out["source_file"] = underlying_path.name

    return out[out_cols]


def parse_citics_derivative_holdings(path: Path) -> pd.DataFrame:
    """解析中信证券代发的 CICC 场外衍生品估值表 .xlsx。

    提取 "互换标的信息" sheet 中的底层标的持仓，按 ticker 去重合并。

    返回列：
        - ticker: 标的代码（如 TSLA.OQ → TSLA）
        - company: 标的名称
        - shares: 合约数量（取各合约数量之和）
        - market_value_cny: 市值（人民币，取最后一列 "市值(人民币)"）
        - source_file: 来源文件名
    """
    xl = pd.ExcelFile(path)
    if len(xl.sheet_names) < 3:
        raise ValueError(f"CITICS 衍生品估值表缺少必要 sheet: {path}")

    # Sheet 2 = 互换标的信息
    df_raw = pd.read_excel(path, sheet_name=2, header=0, dtype=object, engine="openpyxl")
    if df_raw.empty:
        raise ValueError(f"CITICS 衍生品估值表互换标的信息为空: {path}")

    # 过滤有效数据行：标的代码列（第3列，索引2）包含 '.' 且长度 > 3
    rows: list[dict[str, Any]] = []
    for _, row in df_raw.iterrows():
        code = str(row.iloc[2] or "").strip()
        name = str(row.iloc[1] or "").strip()
        if not code or "." not in code or len(code) <= 3:
            continue

        qty = to_float(row.iloc[4])  # 合约数量
        mv = to_float(row.iloc[-1])  # 市值(人民币)
        if mv is None:
            continue

        # 清理 ticker：去掉后缀（如 .OQ, .HK, .N）
        ticker = code.split(".")[0] if "." in code else code

        rows.append({
            "ticker": ticker,
            "company": name,
            "shares": int(qty) if qty is not None else 0,
            "market_value_cny": mv,
        })

    if not rows:
        return pd.DataFrame(columns=["ticker", "company", "shares", "market_value_cny", "source_file"])

    df = pd.DataFrame(rows)
    # 按 ticker 合并（同一个股票可能在多个合约中出现）
    aggregated = df.groupby("ticker", as_index=False).agg(
        company=("company", "first"),
        shares=("shares", "sum"),
        market_value_cny=("market_value_cny", "sum"),
    )
    aggregated["source_file"] = path.name
    return aggregated
