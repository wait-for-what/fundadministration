"""SQLAlchemy engine 工厂。

用途:
- 基于 FUND_DB_URL 创建 engine；支持 sqlite:/// 与 mysql+pymysql:// 两种 URL。

输入:
- 环境变量 FUND_DB_URL，或显式传入 db_url。

输出:
- SQLAlchemy Engine（pool_pre_ping=True, future=True）。

失败行为:
- 未提供 db_url 且 FUND_DB_URL 未设置时抛 RuntimeError。
"""

from __future__ import annotations

from typing import Optional

from sqlalchemy import create_engine, event
from sqlalchemy.engine import Engine

from fundadmin.core.config import get_env


def get_engine(db_url: Optional[str] = None, **kwargs) -> Engine:
    """创建 engine；缺省读取 FUND_DB_URL。

    对 sqlite 启用 WAL + busy_timeout：常驻的 Streamlit 看板（只读）需要与
    18:00 同步任务（写）并发访问同一个 .db 文件。WAL 允许读写并发，
    busy_timeout 让短暂的锁竞争"阻塞重试"而非立刻抛 ``database is locked``
    导致静默丢写/脏读（reliability-3）。
    """
    if not db_url:
        db_url = get_env("FUND_DB_URL", required=True)

    is_sqlite = db_url.startswith("sqlite")
    if is_sqlite:
        connect_args = dict(kwargs.pop("connect_args", {}))
        connect_args.setdefault("timeout", 30)
        kwargs["connect_args"] = connect_args

    engine = create_engine(db_url, pool_pre_ping=True, future=True, **kwargs)

    if is_sqlite:

        @event.listens_for(engine, "connect")
        def _set_sqlite_pragma(dbapi_conn, _record):  # noqa: ANN001
            cur = dbapi_conn.cursor()
            try:
                cur.execute("PRAGMA journal_mode=WAL")
                cur.execute("PRAGMA busy_timeout=30000")
                cur.execute("PRAGMA synchronous=NORMAL")
            finally:
                cur.close()

    return engine


__all__ = ["get_engine"]
