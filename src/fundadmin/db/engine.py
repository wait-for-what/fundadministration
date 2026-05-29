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

from sqlalchemy import create_engine
from sqlalchemy.engine import Engine

from fundadmin.core.config import get_env


def get_engine(db_url: Optional[str] = None, **kwargs) -> Engine:
    """创建 engine；缺省读取 FUND_DB_URL。"""
    if not db_url:
        db_url = get_env("FUND_DB_URL", required=True)
    return create_engine(db_url, pool_pre_ping=True, future=True, **kwargs)


__all__ = ["get_engine"]
