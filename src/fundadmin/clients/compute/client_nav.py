"""客户净值合并。

用途:
- 把 clients 表（custname, prodcode, holding_shares, email）与
  product_nav_history 在指定估值日连接，得到每个客户每个产品的市值。

输入:
- SQLAlchemy Engine、可选 as_of_date（缺省取每个 prodcode 的最新净值）。

输出:
- DataFrame，列：custname, prodcode, prodname, holding_shares, nav_unit,
  nav_cum, market_value, email, mobile, as_of_date

失败行为:
- 客户表为空或净值表为空时返回空 DataFrame。
"""

from __future__ import annotations

import pandas as pd
from sqlalchemy.engine import Engine

from ..config import PRODCODE_TO_NAME
from ..store import load_clients, load_latest_nav, load_nav_asof

_OUTPUT_COLS: tuple[str, ...] = (
    "custname",
    "prodcode",
    "prodname",
    "holding_shares",
    "nav_unit",
    "nav_cum",
    "market_value",
    "email",
    "mobile",
    "as_of_date",
)


def build_client_nav_frame(
    *,
    as_of_date: str | None = None,
    engine: Engine | None = None,
    active_only: bool = True,
) -> pd.DataFrame:
    """连接客户与产品净值，估算市值。"""
    clients = load_clients(active_only=active_only, engine=engine)
    if clients.empty:
        return pd.DataFrame(columns=list(_OUTPUT_COLS))

    if as_of_date:
        nav = load_nav_asof(as_of_date, engine=engine)
    else:
        nav = load_latest_nav(engine)

    if nav.empty:
        nav = pd.DataFrame(columns=["prodcode", "as_of_date", "nav_unit", "nav_cum"])

    merged = clients.merge(
        nav[["prodcode", "as_of_date", "nav_unit", "nav_cum"]],
        on="prodcode",
        how="left",
    )

    if "prodname" in merged.columns:
        merged["prodname"] = merged["prodname"].where(
            merged["prodname"].astype(str).str.strip().ne(""),
            merged["prodcode"].map(PRODCODE_TO_NAME),
        )
    else:
        merged["prodname"] = merged["prodcode"].map(PRODCODE_TO_NAME)

    merged["holding_shares"] = pd.to_numeric(merged["holding_shares"], errors="coerce")
    merged["nav_unit"] = pd.to_numeric(merged["nav_unit"], errors="coerce")
    merged["nav_cum"] = pd.to_numeric(merged["nav_cum"], errors="coerce")
    merged["market_value"] = (merged["holding_shares"] * merged["nav_unit"]).round(2)

    for col in _OUTPUT_COLS:
        if col not in merged.columns:
            merged[col] = None

    return merged[list(_OUTPUT_COLS)].sort_values(["custname", "prodcode"]).reset_index(drop=True)


__all__ = ["build_client_nav_frame"]
