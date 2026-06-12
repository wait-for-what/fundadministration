"""内网看板表的读写适配。

用途:
- 提供按主键覆盖式 upsert（ON CONFLICT DO UPDATE）的写入函数。
- 提供常用的查询函数：最新阈值版本、最新价、产品最近一次估值、客户列表等。

输入:
- SQLAlchemy Engine（默认 schema.get_engine()）。
- pandas DataFrame，列名需与 schema 一致。

输出:
- upsert_*: 返回写入行数。
- load_*:   返回 DataFrame。

失败行为:
- 列缺失时抛 KeyError；类型不匹配时由 SQLAlchemy/SQLite 报错。
"""

from __future__ import annotations

import json
from collections.abc import Mapping, Sequence

import pandas as pd
from sqlalchemy import text
from sqlalchemy.engine import Engine

from .schema import get_engine

# ---------------------------------------------------------------------------
# 内部辅助
# ---------------------------------------------------------------------------

_HOLDINGS_COLS: tuple[str, ...] = (
    "as_of_date",
    "product_code",
    "product_name",
    "broker",
    "asset_class",
    "ticker",
    "instrument_name",
    "market_value",
    "weight",
    "quantity",
    "raw_payload",
)
_CLIENTS_COLS: tuple[str, ...] = (
    "custname",
    "prodcode",
    "prodname",
    "holding_shares",
    "email",
    "mobile",
    "active",
)
_NAV_COLS: tuple[str, ...] = ("prodcode", "as_of_date", "nav_unit", "nav_cum", "src_xlsx")
_ATTACHMENT_COLS: tuple[str, ...] = (
    "sha256",
    "file_name",
    "file_suffix",
    "broker",
    "product_code",
    "as_of_date",
    "attachment_type",
    "sheet_count",
    "row_count",
    "inbox_dir",
)
_RAW_ROW_COLS: tuple[str, ...] = (
    "ingest_id",
    "sheet_index",
    "sheet_name",
    "row_index",
    "cells_json",
)
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
_TRANSACTION_COLS: tuple[str, ...] = (
    "trade_date",
    "product_code",
    "product_name",
    "broker",
    "ticker",
    "instrument_name",
    "direction",
    "open_close",
    "price_local",
    "quantity",
    "amount_local",
    "amount_cny",
    "fee",
    "realized_pnl",
    "currency",
    "contract_no",
    "occ",
    "source_file",
    "ingest_id",
)
# 成交流水主键：occ 为文件内同键去重序号（区分真实重复成交）。
_TRANSACTION_PK: tuple[str, ...] = (
    "broker",
    "product_code",
    "ticker",
    "trade_date",
    "direction",
    "open_close",
    "price_local",
    "quantity",
    "contract_no",
    "occ",
)
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


def _normalize(df: pd.DataFrame, cols: Sequence[str], *, defaults: Mapping[str, object] | None = None) -> pd.DataFrame:
    """保留指定列；缺失列按默认值补齐。"""
    out = df.copy()
    defaults = dict(defaults or {})
    for col in cols:
        if col not in out.columns:
            out[col] = defaults.get(col)
    return out[list(cols)]


def _upsert(engine: Engine, table: str, cols: Sequence[str], pk_cols: Sequence[str], df: pd.DataFrame) -> int:
    """通用 SQLite UPSERT。

    - SQLite 3.24+ 支持 ON CONFLICT DO UPDATE；本仓库环境（Python 3.10+）默认满足。
    - DataFrame 行通过 executemany 一次提交，保留事务原子性。
    """
    if df.empty:
        return 0

    set_cols = [c for c in cols if c not in pk_cols]
    placeholders = ", ".join(f":{c}" for c in cols)
    cols_sql = ", ".join(cols)
    set_sql = ", ".join(f"{c} = excluded.{c}" for c in set_cols) or f"{pk_cols[0]} = excluded.{pk_cols[0]}"
    pk_sql = ", ".join(pk_cols)
    sql = text(
        f"INSERT INTO {table} ({cols_sql}) VALUES ({placeholders}) "
        f"ON CONFLICT({pk_sql}) DO UPDATE SET {set_sql}"
    )

    rows = df.where(pd.notna(df), None).to_dict(orient="records")
    with engine.begin() as conn:
        conn.execute(sql, rows)
    return len(rows)


# ---------------------------------------------------------------------------
# upsert
# ---------------------------------------------------------------------------


def upsert_holdings(df: pd.DataFrame, *, engine: Engine | None = None) -> int:
    """写入产品持仓。主键 (as_of_date, product_code, broker, ticker, instrument_name)。"""
    eng = engine or get_engine()
    payload = _normalize(df, _HOLDINGS_COLS, defaults={"broker": "", "ticker": "", "instrument_name": ""})
    for col in ("broker", "ticker", "instrument_name"):
        payload[col] = payload[col].fillna("").astype(str)
    if "raw_payload" in payload.columns:
        payload["raw_payload"] = payload["raw_payload"].apply(
            lambda v: v if (v is None or isinstance(v, str)) else json.dumps(v, ensure_ascii=False)
        )
    return _upsert(
        eng,
        "fund_portfolio_holdings",
        _HOLDINGS_COLS,
        ("as_of_date", "product_code", "broker", "ticker", "instrument_name"),
        payload,
    )


def upsert_clients(df: pd.DataFrame, *, engine: Engine | None = None) -> int:
    """写入客户主表。主键 (custname, prodcode)。"""
    eng = engine or get_engine()
    payload = _normalize(df, _CLIENTS_COLS, defaults={"active": 1})
    payload["active"] = pd.to_numeric(payload["active"], errors="coerce").fillna(1).astype(int)
    return _upsert(eng, "clients", _CLIENTS_COLS, ("custname", "prodcode"), payload)


def upsert_nav(df: pd.DataFrame, *, engine: Engine | None = None) -> int:
    """写入产品净值。主键 (prodcode, as_of_date)。"""
    eng = engine or get_engine()
    payload = _normalize(df, _NAV_COLS)
    return _upsert(eng, "product_nav_history", _NAV_COLS, ("prodcode", "as_of_date"), payload)


# ---------------------------------------------------------------------------
# 附件入库：原始落地层
# ---------------------------------------------------------------------------


def insert_attachment(meta: Mapping[str, object], *, engine: Engine | None = None) -> tuple[int, bool]:
    """登记一个附件（原始落地层），按内容 sha256 去重。

    返回 (ingest_id, is_new)：
    - sha256 已存在时返回旧 ingest_id 与 is_new=False（调用方应跳过原始行写入）。
    - 否则插入新行返回新 ingest_id 与 is_new=True。
    """
    eng = engine or get_engine()
    sha = str(meta.get("sha256") or "").strip()
    if not sha:
        raise ValueError("insert_attachment: sha256 is empty")

    with eng.begin() as conn:
        existing = conn.execute(
            text("SELECT ingest_id FROM attachment_ingest WHERE sha256 = :s"),
            {"s": sha},
        ).fetchone()
        if existing is not None:
            return int(existing[0]), False

        cols_sql = ", ".join(_ATTACHMENT_COLS)
        placeholders = ", ".join(f":{c}" for c in _ATTACHMENT_COLS)
        params = {c: meta.get(c) for c in _ATTACHMENT_COLS}
        result = conn.execute(
            text(f"INSERT INTO attachment_ingest ({cols_sql}) VALUES ({placeholders})"),
            params,
        )
        new_id = int(result.lastrowid)
    return new_id, True


def upsert_raw_sheet_rows(df: pd.DataFrame, *, engine: Engine | None = None) -> int:
    """写入附件原始行。主键 (ingest_id, sheet_index, row_index)。"""
    eng = engine or get_engine()
    payload = _normalize(df, _RAW_ROW_COLS)
    return _upsert(
        eng,
        "raw_sheet_rows",
        _RAW_ROW_COLS,
        ("ingest_id", "sheet_index", "row_index"),
        payload,
    )


# ---------------------------------------------------------------------------
# 附件入库：核心结构层
# ---------------------------------------------------------------------------


def upsert_positions(df: pd.DataFrame, *, engine: Engine | None = None) -> int:
    """写入持仓超集。主键 (as_of_date, product_code, broker, ticker, instrument_name, contract_no)。"""
    eng = engine or get_engine()
    payload = _normalize(df, _POSITION_COLS, defaults={"broker": "", "contract_no": ""})
    for col in ("broker", "ticker", "instrument_name", "contract_no"):
        payload[col] = payload[col].fillna("").astype(str)
    return _upsert(
        eng,
        "fund_positions",
        _POSITION_COLS,
        ("as_of_date", "product_code", "broker", "ticker", "instrument_name", "contract_no"),
        payload,
    )


def _assign_occ(df: pd.DataFrame, pk_no_occ: Sequence[str]) -> pd.Series:
    """为同键的真实重复成交分配出现序号 occ（0,1,2...）。

    pandas groupby 会丢弃含 NaN 的分组键，故用稳定字符串键做 cumcount，
    保证 None/NaN 也能参与分组。全历史在每日报告中按相同顺序重复出现，
    因此同一笔成交在不同文件里获得相同 occ → ON CONFLICT 折叠去重；
    文件内真实重复的同键成交获得 0/1/2 → 不被误折叠。
    """
    def _key(row: pd.Series) -> str:
        parts = []
        for c in pk_no_occ:
            v = row.get(c)
            parts.append("" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))
        return "\x1f".join(parts)

    keys = df.apply(_key, axis=1)
    return keys.groupby(keys).cumcount()


def upsert_transactions(df: pd.DataFrame, *, engine: Engine | None = None) -> int:
    """写入成交流水。主键见 _TRANSACTION_PK（含 occ 去重序号）。

    occ 由本函数按"主键去 occ"分组计算，调用方无需提供。
    """
    eng = engine or get_engine()
    payload = _normalize(
        df,
        _TRANSACTION_COLS,
        defaults={
            "broker": "",
            "product_name": "",
            "ticker": "",
            "instrument_name": "",
            "direction": "",
            "open_close": "",
            "contract_no": "",
        },
    )
    for col in ("broker", "ticker", "instrument_name", "direction", "open_close", "contract_no"):
        payload[col] = payload[col].fillna("").astype(str)
    pk_no_occ = [c for c in _TRANSACTION_PK if c != "occ"]
    payload["occ"] = _assign_occ(payload, pk_no_occ).astype(int)
    return _upsert(eng, "fund_transactions", _TRANSACTION_COLS, _TRANSACTION_PK, payload)


def upsert_product_valuation(df: pd.DataFrame, *, engine: Engine | None = None) -> int:
    """写入产品级净值/规模。主键 (as_of_date, product_code)。"""
    eng = engine or get_engine()
    payload = _normalize(df, _VALUATION_COLS)
    return _upsert(eng, "product_valuation", _VALUATION_COLS, ("as_of_date", "product_code"), payload)


# ---------------------------------------------------------------------------
# load
# ---------------------------------------------------------------------------


def load_holdings(
    *,
    product_code: str | None = None,
    as_of_date: str | None = None,
    engine: Engine | None = None,
) -> pd.DataFrame:
    """按产品/日期筛选持仓；不传则返回全部。"""
    eng = engine or get_engine()
    clauses = []
    params: dict[str, object] = {}
    if product_code:
        clauses.append("product_code = :product_code")
        params["product_code"] = product_code
    if as_of_date:
        clauses.append("as_of_date = :as_of_date")
        params["as_of_date"] = as_of_date
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    sql = text(f"SELECT * FROM fund_portfolio_holdings {where} ORDER BY as_of_date DESC, weight DESC")
    return pd.read_sql(sql, eng, params=params)


def list_holdings_dates(engine: Engine | None = None) -> list[str]:
    """所有 as_of_date 列表（倒序）。"""
    eng = engine or get_engine()
    sql = text("SELECT DISTINCT as_of_date FROM fund_portfolio_holdings ORDER BY as_of_date DESC")
    with eng.connect() as conn:
        return [row[0] for row in conn.execute(sql).fetchall()]


def list_holdings_products(engine: Engine | None = None) -> list[str]:
    """所有 product_code 列表。"""
    eng = engine or get_engine()
    sql = text("SELECT DISTINCT product_code FROM fund_portfolio_holdings ORDER BY product_code")
    with eng.connect() as conn:
        return [row[0] for row in conn.execute(sql).fetchall()]


def load_clients(*, active_only: bool = True, engine: Engine | None = None) -> pd.DataFrame:
    """读取客户主表。"""
    eng = engine or get_engine()
    sql = "SELECT * FROM clients"
    if active_only:
        sql += " WHERE active = 1"
    sql += " ORDER BY custname, prodcode"
    return pd.read_sql(text(sql), eng)


def load_latest_nav(engine: Engine | None = None) -> pd.DataFrame:
    """每个 prodcode 取最大 as_of_date 的一条净值。"""
    eng = engine or get_engine()
    sql = text(
        """
        SELECT n.*
        FROM product_nav_history n
        JOIN (
            SELECT prodcode, MAX(as_of_date) AS max_d FROM product_nav_history GROUP BY prodcode
        ) t ON n.prodcode = t.prodcode AND n.as_of_date = t.max_d
        ORDER BY n.prodcode
        """
    )
    return pd.read_sql(sql, eng)


def load_nav_asof(as_of_date: str, *, engine: Engine | None = None) -> pd.DataFrame:
    """指定估值日的所有产品净值。"""
    eng = engine or get_engine()
    sql = text("SELECT * FROM product_nav_history WHERE as_of_date = :d ORDER BY prodcode")
    return pd.read_sql(sql, eng, params={"d": as_of_date})


def list_nav_dates(engine: Engine | None = None) -> list[str]:
    """所有净值估值日（倒序）。"""
    eng = engine or get_engine()
    sql = text("SELECT DISTINCT as_of_date FROM product_nav_history ORDER BY as_of_date DESC")
    with eng.connect() as conn:
        return [row[0] for row in conn.execute(sql).fetchall()]


def load_positions(
    *,
    product_code: str | None = None,
    as_of_date: str | None = None,
    engine: Engine | None = None,
) -> pd.DataFrame:
    """按产品/日期筛选 fund_positions；不传则返回全部。"""
    eng = engine or get_engine()
    clauses = []
    params: dict[str, object] = {}
    if product_code:
        clauses.append("product_code = :product_code")
        params["product_code"] = product_code
    if as_of_date:
        clauses.append("as_of_date = :as_of_date")
        params["as_of_date"] = as_of_date
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    sql = text(
        f"SELECT * FROM fund_positions {where} ORDER BY as_of_date DESC, market_value_cny DESC"
    )
    return pd.read_sql(sql, eng, params=params)


def load_product_valuation(
    *,
    as_of_date: str | None = None,
    engine: Engine | None = None,
) -> pd.DataFrame:
    """读取 product_valuation；可按估值日筛选。"""
    eng = engine or get_engine()
    if as_of_date:
        sql = text("SELECT * FROM product_valuation WHERE as_of_date = :d ORDER BY product_code")
        return pd.read_sql(sql, eng, params={"d": as_of_date})
    sql = text("SELECT * FROM product_valuation ORDER BY as_of_date DESC, product_code")
    return pd.read_sql(sql, eng)


def load_transactions(
    *,
    product_code: str | None = None,
    ticker: str | None = None,
    trade_date: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    engine: Engine | None = None,
) -> pd.DataFrame:
    """按产品/标的/日期（或日期区间）筛选成交流水；不传则返回全部。"""
    eng = engine or get_engine()
    clauses = []
    params: dict[str, object] = {}
    if product_code:
        clauses.append("product_code = :product_code")
        params["product_code"] = product_code
    if ticker:
        clauses.append("ticker = :ticker")
        params["ticker"] = ticker
    if trade_date:
        clauses.append("trade_date = :trade_date")
        params["trade_date"] = trade_date
    if date_from:
        clauses.append("trade_date >= :date_from")
        params["date_from"] = date_from
    if date_to:
        clauses.append("trade_date <= :date_to")
        params["date_to"] = date_to
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    sql = text(
        f"SELECT * FROM fund_transactions {where} ORDER BY trade_date, ticker, occ"
    )
    return pd.read_sql(sql, eng, params=params)


def load_raw_sheet_rows(ingest_id: int, *, engine: Engine | None = None) -> pd.DataFrame:
    """读取某个附件的全部原始行（按 sheet/row 顺序）。"""
    eng = engine or get_engine()
    sql = text(
        "SELECT * FROM raw_sheet_rows WHERE ingest_id = :i ORDER BY sheet_index, row_index"
    )
    return pd.read_sql(sql, eng, params={"i": int(ingest_id)})


__all__ = [
    "upsert_holdings",
    "upsert_clients",
    "upsert_nav",
    "insert_attachment",
    "upsert_raw_sheet_rows",
    "upsert_positions",
    "upsert_transactions",
    "upsert_product_valuation",
    "load_holdings",
    "list_holdings_dates",
    "list_holdings_products",
    "load_clients",
    "load_latest_nav",
    "load_nav_asof",
    "list_nav_dates",
    "load_positions",
    "load_transactions",
    "load_product_valuation",
    "load_raw_sheet_rows",
]
