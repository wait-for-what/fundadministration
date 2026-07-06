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

    # 会计括号负数：中文券商估值/Balance 表常用 (1234.50) / (500) 表示负数。
    # 历史实现对此返回 None，导致负持仓/负 P&L 被静默丢弃或经 fillna 变成 0（parsers-1）。
    paren = re.match(r"^\((\d+(?:\.\d+)?)\)$", text)
    if paren:
        try:
            return -float(paren.group(1))
        except Exception:
            return None

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
    """将字符串/数值转换为 int（四舍五入），返回 None 表示失败。

    历史实现用 ``int()`` 向零截断：openpyxl 把整数股数存成 999.9999999 这类浮点噪声时
    会少算一股（parsers-5）。改用 round() 消除浮点噪声导致的 off-by-one。
    """
    f = to_float(value)
    if f is None:
        return None
    return int(round(f))


def normalize_date(value: Any) -> str | None:
    """把多种日期表示归一化为 ISO 日期字符串 YYYY-MM-DD。

    支持:
        - 20201023000000 / 20201023 (中金当日交易的紧凑时间戳)
        - 2026-05-05 / 2026/05/05
        - datetime / pandas.Timestamp
    无法解析时返回 None。
    """
    if value is None:
        return None
    # datetime / Timestamp
    if hasattr(value, "strftime") and not isinstance(value, str):
        try:
            return value.strftime("%Y-%m-%d")
        except Exception:
            pass
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "-", "null"}:
        return None
    # 纯数字紧凑格式：取前 8 位 YYYYMMDD
    digits = re.sub(r"\D", "", text)
    if len(digits) >= 8 and digits[:8].isdigit():
        y, m, d = digits[0:4], digits[4:6], digits[6:8]
        if "1900" <= y <= "2999" and "01" <= m <= "12" and "01" <= d <= "31":
            return f"{y}-{m}-{d}"
    # 回退 pandas 解析
    try:
        ts = pd.to_datetime(text, errors="coerce")
        if pd.notna(ts):
            return ts.strftime("%Y-%m-%d")
    except Exception:
        pass
    return None


def read_csv_robust(
    path: Path, *, encoding: str | None = None, skiprows: int | None = None, **kwargs: Any
) -> pd.DataFrame:
    """尝试多种编码读取 CSV，失败时抛出异常。

    中文券商 CSV 常为 GBK/GB18030 编码；硬编码 utf-8-sig 会抛 UnicodeDecodeError
    或解码出乱码导致表头识别失败、整张表静默丢失（parsers-3）。这里按候选编码依次
    尝试，并透传 ``header`` / ``dtype`` 等参数给 pandas（``dtype`` 默认 object）。
    """
    kwargs.setdefault("dtype", object)
    if skiprows is not None:
        kwargs["skiprows"] = skiprows
    encodings = [encoding] if encoding else []
    encodings += ["utf-8-sig", "utf-8", "gbk", "gb2312", "gb18030", "cp1252"]
    last_exc: Exception | None = None
    for enc in encodings:
        if enc is None:
            continue
        try:
            return pd.read_csv(path, encoding=enc, **kwargs)
        except Exception as exc:
            last_exc = exc
    raise RuntimeError(f"无法以候选编码读取 CSV {path}: {last_exc}")
