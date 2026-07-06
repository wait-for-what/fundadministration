"""资产估值表 NAV 提取器。

逻辑：
- 遍历第一列，找到内容为"资产净值"的所在行。
- 在该行向右扫描，匹配"市值（本币）"列，提取数值。
- 返回 float 类型的 NAV。
"""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Any

import pandas as pd

from fundadmin.portfolio.parsers.common import read_csv_robust, to_float

# 可能的表头别名
NAV_ROW_LABELS = {"资产净值", "净资产", "基金资产净值", "期末资产净值"}
NAV_COL_LABELS = {"市值（本币）", "市值(本币)", "市值（本位币）", "市值(本位币)", "市值", "本币市值"}


def _normalize_cell(value: Any) -> str:
    text = str(value or "").strip()
    text = text.replace("\n", "").replace("\r", "").replace("\t", "").replace(" ", "")
    return text


def _column_header_is_percentish(df: pd.DataFrame, row_idx: int, col_idx: int) -> bool:
    """该列在 NAV 行上方是否存在百分比/占比类表头（如 '市值占净值%'）。

    用于兜底取数时排除占比列，避免把 100% 之类的数字误当作资产净值（parsers-2）。
    """
    for r in range(max(0, row_idx - 30), row_idx):
        header = _normalize_cell(df.iloc[r, col_idx])
        if "%" in header or "占净值" in header or "占比" in header:
            return True
    return False


def parse_nav_from_valuation(path: Path, *, sheet: str | int = 0) -> float:
    """从估值表（CSV 或 Excel）中提取 NAV。

    参数:
        path: 估值表文件路径，支持 .csv / .xlsx / .xls。
        sheet: Excel 时指定工作表名或索引，默认第一个。

    返回:
        NAV 数值。

    异常:
        ValueError: 找不到 NAV 行或对应数值。
    """
    suffix = path.suffix.lower()
    if suffix == ".csv":
        df = read_csv_robust(path, header=None)
    else:
        df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object, engine="openpyxl")

    if df.empty:
        raise ValueError(f"估值表为空: {path}")

    # 限制扫描范围，避免全表遍历
    max_rows = min(200, len(df))
    max_cols = min(40, len(df.columns))

    for row_idx in range(max_rows):
        row_values = df.iloc[row_idx].tolist()
        # 扫描前两列寻找 NAV 行标签（兼容"科目代码"/"科目名称"在第一列的情况）
        found = False
        for col_idx in range(min(2, len(row_values))):
            if _normalize_cell(row_values[col_idx]) in NAV_ROW_LABELS:
                found = True
                break
        if not found:
            continue

        # 找到 NAV 行，向右扫描列头
        row_values = df.iloc[row_idx].tolist()
        for col_idx in range(1, max_cols):
            cell = _normalize_cell(row_values[col_idx])
            if cell in NAV_COL_LABELS:
                # 数值可能在同一行同一列，或下一行同一列
                candidates = [
                    df.iloc[row_idx, col_idx],
                    df.iloc[row_idx + 1, col_idx] if row_idx + 1 < len(df) else None,
                ]
                for val in candidates:
                    num = to_float(val)
                    if num is not None:
                        return num

        # 若按列头未命中，尝试取该行最右侧的非空数值（兜底）。
        # CRITICAL（parsers-2）：必须跳过百分比单元格——估值表"资产净值"行最右常是
        # "市值占净值% = 100.00%"，to_float('100.00%') 返回 1.0，会把 NAV 静默替换成 ~1.0，
        # 进而污染全部 unit_nav*shares 与 weight=mv/nav。跳过任何含 '%' 的单元格，
        # 并对该列上方表头含 '%'/'占净值' 的列一并跳过。
        for col_idx in reversed(range(1, max_cols)):
            raw = df.iloc[row_idx, col_idx]
            cell_text = _normalize_cell(raw)
            if not cell_text or "%" in cell_text:
                continue
            if _column_header_is_percentish(df, row_idx, col_idx):
                continue
            num = to_float(raw)
            if num is not None:
                warnings.warn(
                    f"NAV 按已知列头未命中，回退取行内最右非百分比数值（{path.name} 第 {row_idx} 行，"
                    f"第 {col_idx} 列）；建议在 NAV_COL_LABELS 增补该模板列名以避免猜测。",
                    stacklevel=2,
                )
                return num

        raise ValueError(f"在 {path} 中找到'资产净值'行，但无法提取对应数值")

    raise ValueError(f"在 {path} 中未找到'资产净值'行")
