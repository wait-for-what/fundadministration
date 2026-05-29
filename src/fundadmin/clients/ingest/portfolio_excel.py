"""产品持仓导入：从通用模板 Excel 落到 fund_portfolio_holdings。

用途:
- 接受一份合并后的持仓表（每行一个持仓条目），写入 SQLite。
- 模板列：as_of_date, product_code, product_name, broker, asset_class,
  ticker, instrument_name, market_value, weight, quantity（其余列原样写到 raw_payload）。
- 各券商原始 Excel 仍然由 fund_portfolio.parsers 处理，本函数只负责入库。

输入:
- xlsx: 通用模板 Excel 路径。
- as_of_override: 当 Excel 中 as_of_date 为空时使用。

输出:
- 实际写入的行数。

失败行为:
- 必需列缺失或 product_code 为空时抛 ValueError。
"""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pandas as pd
from sqlalchemy.engine import Engine

from ..config import PRODCODE_TO_NAME
from ..store import upsert_holdings

REQUIRED_COLS = ("product_code",)
OPTIONAL_COLS = (
    "as_of_date",
    "product_name",
    "broker",
    "asset_class",
    "ticker",
    "instrument_name",
    "market_value",
    "weight",
    "quantity",
)


def import_portfolio_xlsx(
    xlsx: str | Path,
    *,
    as_of_override: str | None = None,
    engine: Engine | None = None,
) -> int:
    """从通用模板导入持仓。"""
    file_path = Path(xlsx).expanduser().resolve()
    if not file_path.exists():
        raise FileNotFoundError(file_path)

    df = pd.read_excel(file_path, dtype=object, engine="openpyxl")
    if df.empty:
        return 0
    df = df.rename(columns={c: str(c).strip() for c in df.columns})

    for col in REQUIRED_COLS:
        if col not in df.columns:
            raise ValueError(f"持仓模板缺少必要列: {col}")

    df["product_code"] = df["product_code"].astype(str).str.strip()
    df = df[df["product_code"].astype(bool)]
    if df.empty:
        return 0

    for col in OPTIONAL_COLS:
        if col not in df.columns:
            df[col] = None

    if as_of_override:
        df["as_of_date"] = df["as_of_date"].fillna(as_of_override)
    df["as_of_date"] = df["as_of_date"].fillna(date.today().isoformat())
    df["as_of_date"] = df["as_of_date"].apply(lambda v: pd.to_datetime(v).date().isoformat())

    df["product_name"] = df["product_name"].where(
        df["product_name"].astype(str).str.strip().ne("") & df["product_name"].notna(),
        df["product_code"].map(PRODCODE_TO_NAME),
    )
    df["product_name"] = df["product_name"].fillna(df["product_code"])

    for num_col in ("market_value", "weight", "quantity"):
        df[num_col] = pd.to_numeric(df[num_col], errors="coerce")

    extra_cols = [c for c in df.columns if c not in (*REQUIRED_COLS, *OPTIONAL_COLS, "raw_payload")]
    if extra_cols:
        df["raw_payload"] = df[extra_cols].apply(
            lambda row: json.dumps({k: (None if pd.isna(v) else v) for k, v in row.items()}, ensure_ascii=False, default=str),
            axis=1,
        )
    else:
        df["raw_payload"] = None

    return upsert_holdings(df, engine=engine)


__all__ = ["import_portfolio_xlsx"]
