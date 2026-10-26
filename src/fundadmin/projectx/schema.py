"""projectx 库表结构定义与初始化（MySQL 为主方言）。

用途:
- 在 projectx（MySQL 8，本机 127.0.0.1:3306）中创建基金两张核心表：
  fund_positions（分券商持仓明细）、fund_product_valuation（产品级净值/规模）。
- 与本地 SQLite 的 fund_positions / product_valuation 列一一对应，命名加 fund_ 前缀
  以便在共享库中命名空间隔离（projectx 已有 x_cust_rpt / etf_rotation_universe）。
- 附带 SQLite 方言变体：同步/校验逻辑与方言无关，测试用 SQLite 走同一套代码。

输入:
- PROJECTX_DB_URL（mysql+pymysql://.../projectx?charset=utf8mb4）；
  缺省时回退 DB_URL 并把 schema 改写为 projectx（与 scripts/export_x_cust_rpt.py 一致）。

输出:
- get_engine() 返回 projectx Engine。
- init_db() 建全部表与索引；可重复执行幂等。

失败行为:
- 缺连接串时抛 RuntimeError；建表冲突由 SQLAlchemy/MySQL 报错。
"""

from __future__ import annotations

from sqlalchemy import text
from sqlalchemy.engine import Engine

from fundadmin.core.config import get_env
from fundadmin.db.engine import get_engine as _create_engine

# ---------------------------------------------------------------------------
# MySQL 方言 DDL（目标：MySQL 8.x，utf8mb4）
# ---------------------------------------------------------------------------

_MYSQL_DDL: tuple[str, ...] = (
    """
    CREATE TABLE IF NOT EXISTS fund_positions (
        as_of_date         DATE         NOT NULL COMMENT '估值日',
        product_code       VARCHAR(32)  NOT NULL COMMENT '产品代码',
        product_name       VARCHAR(128) NOT NULL COMMENT '产品名称',
        broker             VARCHAR(32)  NOT NULL DEFAULT '' COMMENT '券商来源(cicc/citic/swhysc)',
        ticker             VARCHAR(64)  NOT NULL DEFAULT '' COMMENT '标的代码',
        instrument_name    VARCHAR(255) NOT NULL DEFAULT '' COMMENT '标的名称',
        asset_class        VARCHAR(32)  NULL COMMENT '资产类别',
        direction          VARCHAR(16)  NULL COMMENT '方向(多/空)',
        currency           VARCHAR(8)   NULL COMMENT '币种',
        fx_rate            DOUBLE       NULL COMMENT '汇率',
        quantity           DOUBLE       NULL COMMENT '数量',
        market_value_cny   DOUBLE       NULL COMMENT '市值(人民币)',
        market_value_local DOUBLE       NULL COMMENT '市值(原币)',
        weight             DOUBLE       NULL COMMENT '占净值比',
        cost_price_local   DOUBLE       NULL COMMENT '成本价(原币)',
        cost_value_local   DOUBLE       NULL COMMENT '成本值(原币)',
        cost_ccy           VARCHAR(8)   NULL COMMENT '成本币种',
        financing_cost     DOUBLE       NULL COMMENT '融资成本',
        init_margin        DOUBLE       NULL COMMENT '初始保证金',
        maint_margin       DOUBLE       NULL COMMENT '维持保证金',
        mtm_pnl            DOUBLE       NULL COMMENT '当日盈亏',
        contract_no        VARCHAR(64)  NOT NULL DEFAULT '' COMMENT '合约号',
        source_files       TEXT         NULL COMMENT '来源附件(逗号分隔)',
        ingest_id          BIGINT       NULL COMMENT '本地附件入库ID',
        updated_at         DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '同步时间',
        PRIMARY KEY (as_of_date, product_code, broker, ticker, instrument_name, contract_no),
        KEY ix_fund_positions_product (product_code, as_of_date)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
    """,
    """
    CREATE TABLE IF NOT EXISTS fund_product_valuation (
        as_of_date             DATE         NOT NULL COMMENT '估值日',
        product_code           VARCHAR(32)  NOT NULL COMMENT '产品代码',
        product_name           VARCHAR(128) NOT NULL COMMENT '产品名称',
        unit_nav               DOUBLE       NULL COMMENT '单位净值',
        asset_nav              DOUBLE       NULL COMMENT '资产净值',
        nav_for_weight         DOUBLE       NULL COMMENT '计权重用净值',
        total_holdings         INT          NULL COMMENT '持仓标的数',
        total_market_value_cny DOUBLE       NULL COMMENT '持仓总市值(人民币)',
        ingest_id              BIGINT       NULL COMMENT '本地附件入库ID',
        updated_at             DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '同步时间',
        PRIMARY KEY (as_of_date, product_code),
        KEY ix_fund_product_valuation_code (product_code, as_of_date)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci
    """,
)

# ---------------------------------------------------------------------------
# SQLite 方言 DDL（仅用于测试：同步/校验代码与方言无关，用 SQLite 走全流程）
# ---------------------------------------------------------------------------

_SQLITE_DDL: tuple[str, ...] = (
    """
    CREATE TABLE IF NOT EXISTS fund_positions (
        as_of_date         TEXT NOT NULL,
        product_code       TEXT NOT NULL,
        product_name       TEXT NOT NULL,
        broker             TEXT NOT NULL DEFAULT '',
        ticker             TEXT NOT NULL DEFAULT '',
        instrument_name    TEXT NOT NULL DEFAULT '',
        asset_class        TEXT,
        direction          TEXT,
        currency           TEXT,
        fx_rate            REAL,
        quantity           REAL,
        market_value_cny   REAL,
        market_value_local REAL,
        weight             REAL,
        cost_price_local   REAL,
        cost_value_local   REAL,
        cost_ccy           TEXT,
        financing_cost     REAL,
        init_margin        REAL,
        maint_margin       REAL,
        mtm_pnl            REAL,
        contract_no        TEXT NOT NULL DEFAULT '',
        source_files       TEXT,
        ingest_id          INTEGER,
        updated_at         TEXT NOT NULL DEFAULT (datetime('now')),
        PRIMARY KEY (as_of_date, product_code, broker, ticker, instrument_name, contract_no)
    )
    """,
    "CREATE INDEX IF NOT EXISTS ix_fund_positions_product ON fund_positions(product_code, as_of_date)",
    """
    CREATE TABLE IF NOT EXISTS fund_product_valuation (
        as_of_date             TEXT NOT NULL,
        product_code           TEXT NOT NULL,
        product_name           TEXT NOT NULL,
        unit_nav               REAL,
        asset_nav              REAL,
        nav_for_weight         REAL,
        total_holdings         INTEGER,
        total_market_value_cny REAL,
        ingest_id              INTEGER,
        updated_at             TEXT NOT NULL DEFAULT (datetime('now')),
        PRIMARY KEY (as_of_date, product_code)
    )
    """,
    "CREATE INDEX IF NOT EXISTS ix_fund_product_valuation_code ON fund_product_valuation(product_code, as_of_date)",
)

TABLES: tuple[str, ...] = ("fund_positions", "fund_product_valuation")


def resolve_projectx_url() -> str:
    """解析 projectx 连接串。

    优先 PROJECTX_DB_URL；否则回退 DB_URL 并把末段 schema 改写为 projectx
    （与 scripts/export_x_cust_rpt.py 的约定一致）。
    """
    url = get_env("PROJECTX_DB_URL")
    if url:
        return url
    raw = get_env("DB_URL")
    if not raw:
        raise RuntimeError(
            "projectx 连接串缺失：请设置 PROJECTX_DB_URL"
            "（mysql+pymysql://user:pwd@host:3306/projectx?charset=utf8mb4）"
        )
    raw = raw.replace("charset=utf8mb4f", "charset=utf8mb4")
    base = raw.rsplit("/", 1)[0]
    return f"{base}/projectx?charset=utf8mb4"


def get_engine() -> Engine:
    """构造 projectx Engine。"""
    return _create_engine(resolve_projectx_url())


def init_db(engine: Engine | None = None) -> Engine:
    """创建全部表与索引；按引擎方言选 DDL，可重复执行幂等。"""
    eng = engine or get_engine()
    ddl = _MYSQL_DDL if eng.dialect.name == "mysql" else _SQLITE_DDL
    with eng.begin() as conn:
        for stmt in ddl:
            conn.exec_driver_sql(stmt)
    return eng


def drop_db(engine: Engine | None = None) -> None:
    """删除全部基金表（仅测试/重建用）。"""
    eng = engine or get_engine()
    with eng.begin() as conn:
        for table in TABLES:
            conn.execute(text(f"DROP TABLE IF EXISTS {table}"))


__all__ = [
    "TABLES",
    "drop_db",
    "get_engine",
    "init_db",
    "resolve_projectx_url",
]
