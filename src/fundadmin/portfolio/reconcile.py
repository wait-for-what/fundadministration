"""持仓与成交流水的交叉核对（reconciliation）。

核心思想：
- 成交流水（fund_transactions）为全历史逐笔成交。
- 任一估值日 D 的持仓数量，应等于该 ticker 截至 D（含）的累计有向成交净额。
- 因此可用成交净额交叉验证 fund_positions 的持仓快照，并定位差异。

提供：
- reconcile_positions_vs_trades(): 对指定产品的若干估值日做核对，返回差异明细。
- reconcile_around_date(): 以某估值日为中心，连同前后各 ±2 个估值日一并核对，
  对应需求"可以和前后两个交易日的持仓量进行交叉比较"。
"""

from __future__ import annotations

from typing import Any

import pandas as pd
from sqlalchemy.engine import Engine

from fundadmin.clients.store import (
    get_engine,
    load_positions,
    load_transactions,
)
from fundadmin.portfolio.parsers.common import clean_ticker
from fundadmin.portfolio.parsers.trades import _signed_qty

RECON_COLS = [
    "product_code",
    "as_of_date",
    "broker",
    "ticker",
    "holding_qty",
    "net_traded_qty",
    "diff",
    "status",
]

# 浮点容差：合约/股数四舍五入误差。
_TOL = 1e-6


def _cumulative_net_qty(trades: pd.DataFrame, as_of_date: str) -> dict[str, float]:
    """计算截至 as_of_date（含）每个 ticker 的累计有向成交净额。"""
    net: dict[str, float] = {}
    if trades is None or trades.empty:
        return net
    sub = trades[trades["trade_date"] <= as_of_date]
    for _, r in sub.iterrows():
        ticker = clean_ticker(r.get("ticker"))
        if not ticker:
            continue
        qty = r.get("quantity")
        try:
            qty = float(qty)
        except (TypeError, ValueError):
            continue
        signed = _signed_qty(
            str(r.get("direction") or ""),
            str(r.get("open_close") or ""),
            qty,
        )
        net[ticker] = net.get(ticker, 0.0) + signed
    return net


def reconcile_positions_vs_trades(
    product_code: str,
    *,
    as_of_dates: list[str] | None = None,
    engine: Engine | None = None,
) -> pd.DataFrame:
    """核对某产品在若干估值日的持仓数量 vs 成交累计净额。

    参数:
        product_code: 产品代码。
        as_of_dates: 待核对的估值日列表；None 表示用 fund_positions 中该产品的全部估值日。

    返回列见 RECON_COLS。status ∈ {ok, mismatch, missing_trades, missing_holding}。
    仅纳入有持仓或有成交净额（任一非零）的 ticker。
    """
    eng = engine or get_engine()
    positions = load_positions(product_code=product_code, engine=eng)
    trades = load_transactions(product_code=product_code, engine=eng)

    if as_of_dates is None:
        if positions.empty:
            return pd.DataFrame(columns=RECON_COLS)
        as_of_dates = sorted(positions["as_of_date"].dropna().unique().tolist())

    rows: list[dict[str, Any]] = []
    for as_of in as_of_dates:
        pos_d = positions[positions["as_of_date"] == as_of]
        # 持仓数量（按 ticker 归一并求和，跨 broker 也合并）。
        hold_qty: dict[str, float] = {}
        hold_broker: dict[str, str] = {}
        for _, r in pos_d.iterrows():
            ticker = clean_ticker(r.get("ticker"))
            if not ticker:
                continue
            try:
                q = float(r.get("quantity") or 0.0)
            except (TypeError, ValueError):
                q = 0.0
            hold_qty[ticker] = hold_qty.get(ticker, 0.0) + q
            hold_broker.setdefault(ticker, str(r.get("broker") or ""))

        net = _cumulative_net_qty(trades, as_of)

        for ticker in sorted(set(hold_qty) | set(net)):
            hq = hold_qty.get(ticker)
            nq = net.get(ticker)
            hq_v = hq if hq is not None else 0.0
            nq_v = nq if nq is not None else 0.0
            if abs(hq_v) < _TOL and abs(nq_v) < _TOL:
                continue
            diff = hq_v - nq_v
            if ticker not in net:
                status = "missing_trades"
            elif ticker not in hold_qty:
                status = "missing_holding"
            elif abs(diff) < _TOL:
                status = "ok"
            else:
                status = "mismatch"
            rows.append(
                {
                    "product_code": product_code,
                    "as_of_date": as_of,
                    "broker": hold_broker.get(ticker, ""),
                    "ticker": ticker,
                    "holding_qty": hq,
                    "net_traded_qty": nq,
                    "diff": diff,
                    "status": status,
                }
            )

    return pd.DataFrame(rows, columns=RECON_COLS)


DELTA_COLS = [
    "product_code",
    "prev_date",
    "curr_date",
    "ticker",
    "holding_delta",
    "net_traded_qty",
    "diff",
    "status",
]


def _net_qty_between(trades: pd.DataFrame, lo_date: str, hi_date: str) -> dict[str, float]:
    """计算 (lo_date, hi_date] 区间内每个 ticker 的有向成交净额（不含 lo_date 当日）。"""
    net: dict[str, float] = {}
    if trades is None or trades.empty:
        return net
    sub = trades[(trades["trade_date"] > lo_date) & (trades["trade_date"] <= hi_date)]
    for _, r in sub.iterrows():
        ticker = clean_ticker(r.get("ticker"))
        if not ticker:
            continue
        try:
            qty = float(r.get("quantity"))
        except (TypeError, ValueError):
            continue
        signed = _signed_qty(
            str(r.get("direction") or ""),
            str(r.get("open_close") or ""),
            qty,
        )
        net[ticker] = net.get(ticker, 0.0) + signed
    return net


def reconcile_holding_deltas(
    product_code: str,
    *,
    as_of_dates: list[str] | None = None,
    engine: Engine | None = None,
) -> pd.DataFrame:
    """相邻估值日持仓变动 vs 区间成交净额的交叉核对（不依赖期初余额）。

    对每对相邻估值日 (prev, curr)：
        holding[curr] - holding[prev]  应等于  (prev, curr] 区间内的成交净额。
    这是对"前后交易日持仓量"最稳健的交叉验证，期初持仓未知也能逐日核对。

    返回列见 DELTA_COLS。status ∈ {ok, mismatch}。
    """
    eng = engine or get_engine()
    positions = load_positions(product_code=product_code, engine=eng)
    trades = load_transactions(product_code=product_code, engine=eng)
    if positions.empty:
        return pd.DataFrame(columns=DELTA_COLS)

    if as_of_dates is None:
        dates = sorted(positions["as_of_date"].dropna().unique().tolist())
    else:
        dates = sorted(set(as_of_dates))

    def _hold_map(as_of: str) -> dict[str, float]:
        pos_d = positions[positions["as_of_date"] == as_of]
        m: dict[str, float] = {}
        for _, r in pos_d.iterrows():
            ticker = clean_ticker(r.get("ticker"))
            if not ticker:
                continue
            try:
                q = float(r.get("quantity") or 0.0)
            except (TypeError, ValueError):
                q = 0.0
            m[ticker] = m.get(ticker, 0.0) + q
        return m

    rows: list[dict[str, Any]] = []
    for prev, curr in zip(dates, dates[1:]):
        hp = _hold_map(prev)
        hc = _hold_map(curr)
        net = _net_qty_between(trades, prev, curr)
        for ticker in sorted(set(hp) | set(hc) | set(net)):
            delta = hc.get(ticker, 0.0) - hp.get(ticker, 0.0)
            nq = net.get(ticker, 0.0)
            if abs(delta) < _TOL and abs(nq) < _TOL:
                continue
            diff = delta - nq
            status = "ok" if abs(diff) < _TOL else "mismatch"
            rows.append(
                {
                    "product_code": product_code,
                    "prev_date": prev,
                    "curr_date": curr,
                    "ticker": ticker,
                    "holding_delta": delta,
                    "net_traded_qty": nq,
                    "diff": diff,
                    "status": status,
                }
            )
    return pd.DataFrame(rows, columns=DELTA_COLS)


def reconcile_around_date(
    product_code: str,
    center_date: str,
    *,
    window: int = 2,
    mode: str = "delta",
    engine: Engine | None = None,
) -> pd.DataFrame:
    """以 center_date 为中心，核对其前后各 window 个估值日（默认 ±2）。

    估值日取自 fund_positions 中该产品实际存在的快照日，按日期排序后
    选取中心日及其相邻的前/后 window 个交易日。

    mode:
        - "delta"（默认）：相邻日持仓变动 vs 区间成交净额（稳健，不依赖期初）。
        - "cumulative"：每个快照日的持仓 vs 累计成交净额（需要全历史期初成交）。
    """
    eng = engine or get_engine()
    positions = load_positions(product_code=product_code, engine=eng)
    if positions.empty:
        return pd.DataFrame(columns=RECON_COLS)

    all_dates = sorted(positions["as_of_date"].dropna().unique().tolist())
    if center_date in all_dates:
        idx = all_dates.index(center_date)
    else:
        # center_date 无快照时，取不晚于它的最近一日为中心。
        earlier = [d for d in all_dates if d <= center_date]
        if not earlier:
            idx = 0
        else:
            idx = all_dates.index(earlier[-1])

    lo = max(0, idx - window)
    hi = min(len(all_dates), idx + window + 1)
    window_dates = all_dates[lo:hi]
    if mode == "delta":
        return reconcile_holding_deltas(product_code, as_of_dates=window_dates, engine=eng)
    return reconcile_positions_vs_trades(product_code, as_of_dates=window_dates, engine=eng)


__all__ = [
    "RECON_COLS",
    "DELTA_COLS",
    "reconcile_positions_vs_trades",
    "reconcile_holding_deltas",
    "reconcile_around_date",
]
