"""客户与产品净值导入。

用途:
- import_clients_csv: 从 projectx.x_cust_rpt 一次性导出的 CSV 落 clients 表。
  期望列：custname, prodcode, prodname, 持有份额(或 holding_shares), email, mobile。
- import_nav_xlsx: 从 fund_portfolio_summary_<date>.xlsx 读 product_name + unit_nav，
  反查 PRODCODE_TO_NAME 映射后落 product_nav_history。
  也可直接接受列：prodcode, as_of_date, nav_unit, nav_cum。

输入:
- 文件路径与可选估值日。

输出:
- 写入行数。

失败行为:
- 列缺失或映射失败时抛 ValueError；CSV 编码异常由 read_csv_robust 报错。
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from sqlalchemy.engine import Engine

from fundadmin.portfolio.parsers.common import read_csv_robust

from ..config import PRODCODE_TO_NAME
from ..store import upsert_clients, upsert_nav

_NAME_TO_PRODCODE: dict[str, str] = {v: k for k, v in PRODCODE_TO_NAME.items()}


def import_clients_csv(csv_path: str | Path, *, engine: Engine | None = None) -> int:
    """从导出的 CSV 写入 clients。"""
    path = Path(csv_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    df = read_csv_robust(path)
    if df.empty:
        return 0
    df = df.rename(columns={c: str(c).strip() for c in df.columns})

    rename_map: dict[str, str] = {}
    if "持有份额" in df.columns and "holding_shares" not in df.columns:
        rename_map["持有份额"] = "holding_shares"
    df = df.rename(columns=rename_map)

    for col in ("custname", "prodcode"):
        if col not in df.columns:
            raise ValueError(f"clients CSV 缺少必要列: {col}")
    for col in ("prodname", "holding_shares", "email", "mobile", "active"):
        if col not in df.columns:
            df[col] = None

    df["custname"] = df["custname"].astype(str).str.strip()
    df["prodcode"] = df["prodcode"].astype(str).str.strip()
    df = df[df["custname"].astype(bool) & df["prodcode"].astype(bool)]
    if df.empty:
        return 0

    df["holding_shares"] = pd.to_numeric(df["holding_shares"], errors="coerce")
    df["active"] = pd.to_numeric(df["active"], errors="coerce").fillna(1).astype(int)
    df["prodname"] = df["prodname"].where(
        df["prodname"].astype(str).str.strip().ne("") & df["prodname"].notna(),
        df["prodcode"].map(PRODCODE_TO_NAME),
    )

    return upsert_clients(df, engine=engine)


def import_nav_xlsx(
    xlsx: str | Path,
    *,
    as_of_date: str | None = None,
    engine: Engine | None = None,
) -> int:
    """从 summary Excel 写入 product_nav_history。"""
    path = Path(xlsx).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    df = pd.read_excel(path, engine="openpyxl")
    if df.empty:
        return 0
    df = df.rename(columns={c: str(c).strip() for c in df.columns})

    if "prodcode" not in df.columns:
        if "product_name" in df.columns:
            df["prodcode"] = df["product_name"].astype(str).str.strip().map(_NAME_TO_PRODCODE)
        else:
            raise ValueError("Excel 缺少 prodcode 或 product_name 列")

    df = df[df["prodcode"].notna() & df["prodcode"].astype(str).str.strip().ne("")]
    if df.empty:
        return 0

    rename_map: dict[str, str] = {}
    if "unit_nav" in df.columns and "nav_unit" not in df.columns:
        rename_map["unit_nav"] = "nav_unit"
    if "asset_nav" in df.columns and "nav_cum" not in df.columns:
        rename_map["asset_nav"] = "nav_cum"
    df = df.rename(columns=rename_map)

    if "nav_unit" not in df.columns:
        raise ValueError("Excel 缺少 nav_unit / unit_nav 列")
    if "as_of_date" not in df.columns:
        df["as_of_date"] = as_of_date or _infer_date_from_filename(path) or date.today().isoformat()
    df["as_of_date"] = df["as_of_date"].apply(lambda v: pd.to_datetime(v).date().isoformat())
    df["nav_unit"] = pd.to_numeric(df["nav_unit"], errors="coerce")
    if "nav_cum" in df.columns:
        df["nav_cum"] = pd.to_numeric(df["nav_cum"], errors="coerce")
    else:
        df["nav_cum"] = None
    df["src_xlsx"] = path.name

    df = df[df["nav_unit"].notna()]
    if df.empty:
        return 0

    return upsert_nav(df, engine=engine)


def _infer_date_from_filename(path: Path) -> str | None:
    """从 fund_portfolio_summary_2026-04-25.xlsx 之类文件名推日期。"""
    import re

    m = re.search(r"(\d{4}-\d{2}-\d{2})", path.stem)
    if not m:
        return None
    try:
        return pd.to_datetime(m.group(1)).date().isoformat()
    except Exception:
        return None


__all__ = ["import_clients_csv", "import_nav_xlsx"]
