"""本地 SQLite -> projectx 的持仓/净值同步与对账。

用途:
- 以本地库（FUND_DB_URL，SQLite）为唯一事实源，把 fund_positions / product_valuation
  按"估值日整快照替换"镜像到 projectx 的 fund_positions / fund_product_valuation：
  每个估值日先 DELETE 再分块 INSERT（单事务），天然幂等，本地删改可完整传播。
- verify(): 按表按估值日对账两侧行数，返回问题清单（空列表 = 一致）。

输入:
- source: 本地 Engine（默认 clients.schema.get_engine()）。
- target: projectx Engine（默认 projectx.schema.get_engine()）。
- dates:  只同步指定估值日（YYYY-MM-DD 列表）；None = 全量镜像。

输出:
- sync_*: 返回 {"table", "dates", "rows"} 统计。
- verify(): 返回问题字符串列表。

失败行为:
- 目标表不存在时抛 SQLAlchemy 异常；先调用 init_db()（sync 命令内已自动建表）。
"""

from __future__ import annotations

import logging
from collections.abc import Sequence

import pandas as pd
from sqlalchemy import text
from sqlalchemy.engine import Engine

from fundadmin.clients.schema import get_engine as _local_engine
from fundadmin.clients.store import load_positions, load_product_valuation

from .schema import get_engine as _projectx_engine

logger = logging.getLogger(__name__)

_CHUNK_SIZE = 500

# 与本地 store._POSITION_COLS 一致（去掉 loaded_at，目标库用 updated_at 记同步时间）。
_POSITION_COLS: tuple[str, ...] = (
    "as_of_date",
    "product_code",
    "product_name",
    "broker",
    "ticker",
    "instrument_name",
    "asset_class",
    "direction",
    "currency",
    "fx_rate",
    "quantity",
    "market_value_cny",
    "market_value_local",
    "weight",
    "cost_price_local",
    "cost_value_local",
    "cost_ccy",
    "financing_cost",
    "init_margin",
    "maint_margin",
    "mtm_pnl",
    "contract_no",
    "source_files",
    "ingest_id",
)
# 与本地 store._VALUATION_COLS 一致（去掉 loaded_at）。
_VALUATION_COLS: tuple[str, ...] = (
    "as_of_date",
    "product_code",
    "product_name",
    "unit_nav",
    "asset_nav",
    "nav_for_weight",
    "total_holdings",
    "total_market_value_cny",
    "ingest_id",
)


def _normalize_dates(raw) -> set[str]:
    """把远端/本地查询出的估值日统一成 'YYYY-MM-DD' 字符串。"""
    return {str(v)[:10] for v in raw}


def _to_rows(df: pd.DataFrame, cols: Sequence[str]) -> list[dict]:
    """DataFrame -> 插入参数行；NaN 统一为 NULL，缺列补 NULL。"""
    payload = df.reindex(columns=list(cols)).copy()
    payload = payload.where(pd.notna(payload), None)
    return payload.to_dict(orient="records")


def _insert_rows(conn, table: str, cols: Sequence[str], rows: list[dict], chunk: int) -> int:
    """分块 executemany 插入，返回写入行数。"""
    if not rows:
        return 0
    cols_sql = ", ".join(cols)
    placeholders = ", ".join(f":{c}" for c in cols)
    sql = text(f"INSERT INTO {table} ({cols_sql}) VALUES ({placeholders})")
    written = 0
    for i in range(0, len(rows), chunk):
        part = rows[i : i + chunk]
        conn.execute(sql, part)
        written += len(part)
    return written


def _remote_dates(conn, table: str) -> set[str]:
    rows = conn.execute(text(f"SELECT DISTINCT as_of_date FROM {table}")).fetchall()
    return _normalize_dates([r[0] for r in rows])


def _mirror_table(
    df: pd.DataFrame,
    *,
    local_table: str,
    target_table: str,
    cols: Sequence[str],
    target: Engine,
    dates: Sequence[str] | None,
    chunk: int,
) -> dict:
    """按估值日整快照镜像一张表（单事务）。"""
    if dates is not None:
        want = set(str(d)[:10] for d in dates)
        df = df[df["as_of_date"].isin(want)]
    plan = sorted(str(d)[:10] for d in df["as_of_date"].unique())

    with target.begin() as conn:
        if dates is None:
            # 全量镜像：本地已不存在的估值日，同步删除远端残留。
            for d in sorted(_remote_dates(conn, target_table) - set(plan)):
                conn.execute(
                    text(f"DELETE FROM {target_table} WHERE as_of_date = :d"), {"d": d}
                )
                logger.info("mirror %s: dropped remote-only date %s", target_table, d)
        for d in plan:
            sub = df[df["as_of_date"].map(lambda v: str(v)[:10]) == d]
            conn.execute(
                text(f"DELETE FROM {target_table} WHERE as_of_date = :d"), {"d": d}
            )
            _insert_rows(conn, target_table, cols, _to_rows(sub, cols), chunk)

    return {"table": target_table, "dates": len(plan), "rows": int(len(df))}


def sync_positions(
    *,
    source: Engine | None = None,
    target: Engine | None = None,
    dates: Sequence[str] | None = None,
    chunk: int = _CHUNK_SIZE,
) -> dict:
    """镜像分券商持仓 fund_positions。"""
    src = source or _local_engine()
    dst = target or _projectx_engine()
    df = load_positions(engine=src)
    return _mirror_table(
        df,
        local_table="fund_positions",
        target_table="fund_positions",
        cols=_POSITION_COLS,
        target=dst,
        dates=dates,
        chunk=chunk,
    )


def sync_valuation(
    *,
    source: Engine | None = None,
    target: Engine | None = None,
    dates: Sequence[str] | None = None,
    chunk: int = _CHUNK_SIZE,
) -> dict:
    """镜像产品级净值/规模 product_valuation -> fund_product_valuation。"""
    src = source or _local_engine()
    dst = target or _projectx_engine()
    df = load_product_valuation(engine=src)
    return _mirror_table(
        df,
        local_table="product_valuation",
        target_table="fund_product_valuation",
        cols=_VALUATION_COLS,
        target=dst,
        dates=dates,
        chunk=chunk,
    )


def sync_all(
    *,
    source: Engine | None = None,
    target: Engine | None = None,
    dates: Sequence[str] | None = None,
    chunk: int = _CHUNK_SIZE,
) -> list[dict]:
    """镜像全部基金表。"""
    return [
        sync_positions(source=source, target=target, dates=dates, chunk=chunk),
        sync_valuation(source=source, target=target, dates=dates, chunk=chunk),
    ]


def _date_counts(engine: Engine, table: str) -> dict[str, int]:
    with engine.connect() as conn:
        rows = conn.execute(
            text(f"SELECT as_of_date, COUNT(*) FROM {table} GROUP BY as_of_date")
        ).fetchall()
    return {str(d)[:10]: int(n) for d, n in rows}


def verify(
    *,
    source: Engine | None = None,
    target: Engine | None = None,
) -> list[str]:
    """对账本地与 projectx：按表按估值日比行数，返回问题清单。"""
    src = source or _local_engine()
    dst = target or _projectx_engine()
    pairs = (
        ("fund_positions", "fund_positions"),
        ("product_valuation", "fund_product_valuation"),
    )
    issues: list[str] = []
    for local_table, remote_table in pairs:
        local_counts = _date_counts(src, local_table)
        remote_counts = _date_counts(dst, remote_table)
        for d in sorted(set(local_counts) - set(remote_counts)):
            issues.append(
                f"{remote_table}: 估值日 {d} 本地 {local_counts[d]} 行，projectx 缺失"
            )
        for d in sorted(set(remote_counts) - set(local_counts)):
            issues.append(
                f"{remote_table}: 估值日 {d} projectx 多出 {remote_counts[d]} 行（本地无）"
            )
        for d in sorted(set(local_counts) & set(remote_counts)):
            if local_counts[d] != remote_counts[d]:
                issues.append(
                    f"{remote_table}: 估值日 {d} 行数不一致 "
                    f"本地={local_counts[d]} projectx={remote_counts[d]}"
                )
        if not local_counts and not remote_counts:
            issues.append(f"{remote_table}: 两侧均为空（未同步或无数据）")
    return issues


__all__ = [
    "sync_all",
    "sync_positions",
    "sync_valuation",
    "verify",
]
