"""一次性数据迁移：旧 internal_dashboard SQLite 的 3 张基金表 -> 新仓 FUND_DB_URL。

用途:
- 剥离切换时，把 MarketAnalysis 时期 internal_dashboard SQLite（INTERNAL_DB_URL）里的
  clients / product_nav_history / fund_portfolio_holdings 三张表，复制到 FundAdministration
  的 FUND_DB_URL（默认独立 SQLite）。幂等：按主键 upsert，可重复执行。

输入:
- --source: 源 SQLite 路径或 SQLAlchemy URL；缺省读 SOURCE_DB_URL / INTERNAL_DB_URL。
- 目标库：FUND_DB_URL（来自 .env）。

输出:
- 打印每张表迁移的行数。

调用示例:
- python scripts/migrate_from_marketanalysis.py --source /path/to/old_internal.db

风险说明:
- 涉及真实投资人 PII，仅在掌握源/目标库访问权限的可信机器上手工执行；不纳入自动化。
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine, text

from fundadmin.clients import store
from fundadmin.clients.schema import init_db

_TABLES: tuple[str, ...] = ("clients", "product_nav_history", "fund_portfolio_holdings")


def _source_engine(source: str):
    s = source.strip()
    if "://" not in s:
        s = f"sqlite:///{Path(s).expanduser().resolve()}"
    return create_engine(s, future=True)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Migrate fund tables from old internal_dashboard SQLite to FUND_DB_URL"
    )
    ap.add_argument(
        "--source",
        default=None,
        help="源 SQLite 路径或 URL；缺省读 SOURCE_DB_URL / INTERNAL_DB_URL",
    )
    args = ap.parse_args()

    source = args.source or os.getenv("SOURCE_DB_URL") or os.getenv("INTERNAL_DB_URL")
    if not source:
        raise SystemExit("缺少源库：请用 --source 或设置 SOURCE_DB_URL / INTERNAL_DB_URL")

    src = _source_engine(source)
    init_db()  # 确保目标 schema 存在

    upserts = {
        "clients": store.upsert_clients,
        "product_nav_history": store.upsert_nav,
        "fund_portfolio_holdings": store.upsert_holdings,
    }
    for table in _TABLES:
        try:
            df = pd.read_sql(text(f"SELECT * FROM {table}"), src)
        except Exception as exc:  # noqa: BLE001 — 源表缺失时跳过并提示
            print(f"[skip] {table}: {exc}")
            continue
        rows = upserts[table](df) if not df.empty else 0
        print(f"[ok] {table}: migrated {rows} rows")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
