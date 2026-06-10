"""SQLite 表结构定义与初始化。

用途:
- 维护本地表：fund_portfolio_holdings / clients / product_nav_history（原有 3 张），
  以及附件入库两层结构 attachment_ingest / raw_sheet_rows（原始无损落地层）、
  fund_positions / product_valuation（核心结构层）。
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
    # ---- Layer 1：原始无损落地层 ----
    # 每个入库附件一行；sha256 为内容去重唯一键，相同内容（即便文件名不同）只入库一次。
    """
    CREATE TABLE IF NOT EXISTS attachment_ingest (
        ingest_id       INTEGER PRIMARY KEY AUTOINCREMENT,
        sha256          TEXT NOT NULL UNIQUE,
        file_name       TEXT NOT NULL,
        file_suffix     TEXT,
        broker          TEXT,
        product_code    TEXT,
        as_of_date      TEXT,
        attachment_type TEXT,
        sheet_count     INTEGER,
        row_count       INTEGER,
        inbox_dir       TEXT,
        ingested_at     TEXT NOT NULL DEFAULT (datetime('now'))
    )
    """,
    "CREATE INDEX IF NOT EXISTS ix_ingest_date_product ON attachment_ingest(as_of_date, product_code)",
    # 每个附件每张表每一行原样存为 JSON 数组（无损），查询用 SQLite json_extract / ->>。
    """
    CREATE TABLE IF NOT EXISTS raw_sheet_rows (
        ingest_id   INTEGER NOT NULL,
        sheet_index INTEGER NOT NULL,
        sheet_name  TEXT,
        row_index   INTEGER NOT NULL,
        cells_json  TEXT NOT NULL,
        PRIMARY KEY (ingest_id, sheet_index, row_index)
    )
    """,
    # ---- Layer 2：核心结构层 ----
    # 持仓超集：现填子集，超集列（保证金/汇率/损益等）留 Phase 2 从原始层回填。
    """
    CREATE TABLE IF NOT EXISTS fund_positions (
        as_of_date          TEXT NOT NULL,
        product_code        TEXT NOT NULL,
        product_name        TEXT NOT NULL,
        broker              TEXT NOT NULL DEFAULT '',
        ticker              TEXT NOT NULL DEFAULT '',
        instrument_name     TEXT NOT NULL DEFAULT '',
        asset_class         TEXT,
        direction           TEXT,
        currency            TEXT,
        fx_rate             REAL,
        quantity            REAL,
        market_value_cny    REAL,
        market_value_local  REAL,
        weight              REAL,
        financing_cost      REAL,
        init_margin         REAL,
        maint_margin        REAL,
        mtm_pnl             REAL,
        contract_no         TEXT NOT NULL DEFAULT '',
        source_files        TEXT,
        ingest_id           INTEGER,
        loaded_at           TEXT NOT NULL DEFAULT (datetime('now')),
        PRIMARY KEY (as_of_date, product_code, broker, ticker, instrument_name, contract_no)
    )
    """,
    "CREATE INDEX IF NOT EXISTS ix_positions_product ON fund_positions(product_code, as_of_date)",
    # 产品级净值/规模（与 product_nav_history 互补，含市值/资产净值）。
    """
    CREATE TABLE IF NOT EXISTS product_valuation (
        as_of_date              TEXT NOT NULL,
        product_code            TEXT NOT NULL,
        product_name            TEXT NOT NULL,
        unit_nav                REAL,
        asset_nav               REAL,
        nav_for_weight          REAL,
        total_holdings          INTEGER,
        total_market_value_cny  REAL,
        ingest_id               INTEGER,
        loaded_at               TEXT NOT NULL DEFAULT (datetime('now')),
        PRIMARY KEY (as_of_date, product_code)
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
