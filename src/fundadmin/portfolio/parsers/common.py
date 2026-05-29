"""通用工具函数。"""

from __future__ import annotations

import re
import warnings
from pathlib import Path
from typing import Any

import pandas as pd

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)


def clean_ticker(code: str | Any) -> str:
    """从带后缀的代码中提取统一 ticker。

    示例:
        - AMZN.US   -> AMZN
        - TSLA.OQ   -> TSLA
        - V.N       -> V
        - 0286.HK   -> 0286
        - 0700.HK   -> 0700
    """
    text = str(code or "").strip().upper()
    # 去掉可能的空格
    text = text.replace(" ", "")
    # 取 . 前面的部分
    if "." in text:
        text = text.split(".")[0]
    return text


def to_float(value: Any) -> float | None:
    """将字符串/数值转换为 float，处理逗号与百分比。

    返回 None 表示无法转换或空值。
    """
    if value is None:
        return None
    if isinstance(value, (int, float)) and pd.notna(value):
        return float(value)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "-", "", "null"}:
        return None

    # 去掉千分位逗号与中文字符逗号
    text = text.replace(",", "").replace("，", "")

    # 百分比
    percent_match = re.match(r"^(-?\d+(?:\.\d+)?)\s*%$", text)
    if percent_match:
        try:
            return float(percent_match.group(1)) / 100.0
        except Exception:
            return None

    try:
        return float(text)
    except Exception:
        return None


def to_int(value: Any) -> int | None:
    """将字符串/数值转换为 int，返回 None 表示失败。"""
    f = to_float(value)
    if f is None:
        return None
    return int(f)


def read_csv_robust(path: Path, *, encoding: str | None = None, skiprows: int | None = None) -> pd.DataFrame:
    """尝试多种编码读取 CSV，失败时抛出异常。"""
    encodings = [encoding] if encoding else []
    encodings += ["utf-8-sig", "utf-8", "gbk", "gb2312", "gb18030", "cp1252"]
    last_exc: Exception | None = None
    for enc in encodings:
        if enc is None:
            continue
        try:
            return pd.read_csv(path, encoding=enc, skiprows=skiprows, dtype=object)
        except Exception as exc:
            last_exc = exc
    raise RuntimeError(f"无法读取 CSV {path}: {last_exc}")
