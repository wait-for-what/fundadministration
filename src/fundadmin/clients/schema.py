"""SQLite 表结构定义与初始化。

用途:
- 维护 3 张本地表：fund_portfolio_holdings / clients / product_nav_history。
- 不使用任何 MySQL 方言（无 ENGINE / utf8mb4 / ON UPDATE CURRENT_TIMESTAMP / AUTO_INCREMENT）。

输入:
- 通过 config.resolve_db_url() 获取 SQLite URL。

输出:
- get_engine() 返回 SQLAlchemy Engine。
- init_db() 创建全部表与索引；可重复执行幂等。
"""

from __future__ import annotations

from sqlalchemy.engine import Engine

from fundadmin.db.engine import get_engine as _create_engine

from .config import resolve_db_url

# 3 张表 + 2 个索引。每条语句独立可重复执行。
_DDL_STATEMENTS: tuple[str, ...] = (
    """
    CREATE TABLE IF NOT EXISTS fund_portfolio_holdings (
        as_of_date      TEXT NOT NULL,
        product_code    TEXT NOT NULL,
        product_name    TEXT NOT NULL,
        broker          TEXT NOT NULL DEFAULT '',
        asset_class     TEXT,
        ticker          TEXT NOT NULL DEFAULT '',
        instrument_name TEXT NOT NULL DEFAULT '',
        market_value    REAL,
        weight          REAL,
        quantity        REAL,
        raw_payload     TEXT,
        loaded_at       TEXT NOT NULL DEFAULT (datetime('now')),
        PRIMARY KEY (as_of_date, product_code, broker, ticker, instrument_name)
    )
    """,
    "CREATE INDEX IF NOT EXISTS ix_holdings_product ON fund_portfolio_holdings(product_code, as_of_date)",
    """
    CREATE TABLE IF NOT EXISTS clients (
        custname        TEXT NOT NULL,
        prodcode        TEXT NOT NULL,
        prodname        TEXT,
        holding_shares  REAL,
        email           TEXT,
        mobile          TEXT,
        active          INTEGER NOT NULL DEFAULT 1,
        updated_at      TEXT NOT NULL DEFAULT (datetime('now')),
        PRIMARY KEY (custname, prodcode)
    )
    """,
    "CREATE INDEX IF NOT EXISTS ix_clients_email ON clients(email)",
    """
    CREATE TABLE IF NOT EXISTS product_nav_history (
        prodcode    TEXT NOT NULL,
        as_of_date  TEXT NOT NULL,
        nav_unit    REAL NOT NULL,
        nav_cum     REAL,
        src_xlsx    TEXT,
        loaded_at   TEXT NOT NULL DEFAULT (datetime('now')),
        PRIMARY KEY (prodcode, as_of_date)
    )
    """,
)


def get_engine() -> Engine:
    """构造看板 SQLite Engine。"""
    return _create_engine(resolve_db_url())


def init_db(engine: Engine | None = None) -> Engine:
    """创建/补齐全部表与索引；多次执行幂等。"""
    eng = engine or get_engine()
    with eng.begin() as conn:
        for stmt in _DDL_STATEMENTS:
            conn.exec_driver_sql(stmt)
    return eng


__all__ = ["get_engine", "init_db"]
