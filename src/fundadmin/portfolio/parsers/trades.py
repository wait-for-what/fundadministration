"""成交流水解析器：中金"当日交易" + 中信"Transaction" sheet。

要点：
- 中金"当日交易"是**全历史**成交（可回溯数年），每天的估值报告都重复携带全量。
  净额（开仓 - 平仓）按 ticker 求和后，等于"持仓"sheet 的合约持仓（已实测逐一吻合）。
- 中信 Statement 的"Transaction"sheet 为逐笔成交（交易方向 B/S），样例中常为空表。

统一输出列：
    trade_date / ticker / instrument_name / direction / open_close /
    price_local / quantity / amount_local / amount_cny / fee /
    realized_pnl / currency / contract_no / source_file
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from fundadmin.portfolio.parsers.common import (
    clean_ticker,
    normalize_date,
    to_float,
)

TX_COLUMNS = [
    "trade_date",
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
    "source_file",
]


def _empty_tx() -> pd.DataFrame:
    return pd.DataFrame(columns=TX_COLUMNS)


def _pick(row: pd.Series, *names: str) -> Any:
    """按列名优先级取第一个存在且非空的值。"""
    for n in names:
        if n in row.index:
            v = row[n]
            if v is not None and str(v).strip().lower() not in {"", "nan", "none"}:
                return v
    return None


def parse_cicc_trades(path: Path, *, sheet: str = "当日交易") -> pd.DataFrame:
    """解析中金估值报告 .xlsx 的"当日交易"sheet（全历史成交）。

    返回标准化成交流水 DataFrame（见模块 docstring）。无该 sheet 时返回空表。
    """
    try:
        df = pd.read_excel(path, sheet_name=sheet, dtype=object, engine="openpyxl")
    except Exception:
        return _empty_tx()
    if df.empty:
        return _empty_tx()
    df.columns = [str(c).strip() for c in df.columns]

    rows: list[dict[str, Any]] = []
    for _, r in df.iterrows():
        code = _pick(r, "标的代码", "证券代码")
        if code is None:
            continue
        qty = to_float(_pick(r, "成交数量", "数量"))
        if qty is None or qty == 0:
            continue
        price = to_float(_pick(r, "成交价格（交易币种）", "成交价格(交易币种)", "成交价格（本币）", "成交价格(本币)"))
        commission = to_float(_pick(r, "佣金")) or 0.0
        stamp = to_float(_pick(r, "印花税")) or 0.0
        rows.append(
            {
                "trade_date": normalize_date(_pick(r, "交易日期")),
                "ticker": clean_ticker(code),
                "instrument_name": str(_pick(r, "标的名称", "证券名称") or "").strip(),
                "direction": str(_pick(r, "成交方向", "交易方向") or "").strip(),
                "open_close": str(_pick(r, "开平仓") or "").strip(),
                "price_local": price,
                "quantity": qty,
                "amount_local": (price * qty) if price is not None else None,
                "amount_cny": to_float(_pick(r, "成交金额（人民币）", "成交金额(人民币)")),
                "fee": commission + stamp,
                "realized_pnl": to_float(_pick(r, "已实现收益")),
                "currency": str(_pick(r, "交易币种", "币种") or "").strip() or None,
                "contract_no": str(_pick(r, "合约号", "合约编号") or "").strip(),
                "source_file": path.name,
            }
        )
    return pd.DataFrame(rows, columns=TX_COLUMNS) if rows else _empty_tx()


def parse_citic_transactions(path: Path, *, sheet: str = "Transaction") -> pd.DataFrame:
    """解析中信 Statement .xlsx 的"Transaction"sheet（逐笔成交）。

    返回标准化成交流水 DataFrame。空表或缺 sheet 时返回空表。
    """
    try:
        df = pd.read_excel(path, sheet_name=sheet, dtype=object, engine="openpyxl")
    except Exception:
        return _empty_tx()
    if df.empty:
        return _empty_tx()
    df.columns = [str(c).strip() for c in df.columns]

    rows: list[dict[str, Any]] = []
    for _, r in df.iterrows():
        code = _pick(r, "证券代码", "标的代码")
        if code is None:
            continue
        qty = to_float(_pick(r, "数量", "成交数量"))
        if qty is None or qty == 0:
            continue
        price = to_float(_pick(r, "价格含费", "价格不含费"))
        rows.append(
            {
                "trade_date": normalize_date(_pick(r, "交易日期")),
                "ticker": clean_ticker(code),
                "instrument_name": str(_pick(r, "证券名称", "标的名称") or "").strip(),
                "direction": str(_pick(r, "交易方向", "成交方向") or "").strip(),
                "open_close": "",
                "price_local": price,
                "quantity": qty,
                "amount_local": to_float(_pick(r, "金额含费", "金额不含费")),
                "amount_cny": None,
                "fee": to_float(_pick(r, "费用")),
                "realized_pnl": to_float(_pick(r, "盈亏含费")),
                "currency": str(_pick(r, "交易货币") or "").strip() or None,
                "contract_no": str(_pick(r, "合约编号", "合约号") or "").strip(),
                "source_file": path.name,
            }
        )
    return pd.DataFrame(rows, columns=TX_COLUMNS) if rows else _empty_tx()


def _signed_qty(direction: str, open_close: str, qty: float) -> float:
    """成交对持仓的有向数量。

    - 中金：开仓为 +，平仓为 -（与"持仓"sheet 合约持仓实测一致）。
    - 中信（无开平仓）：买(B/买)为 +，卖(S/卖)为 -。

    注意：以成交数量的绝对值为基准再赋方向。中信 Transaction 的数量列本身
    已带正负号（卖出为负），若直接对带号数量再取反会双重翻转，故先取绝对值。
    """
    mag = abs(qty)
    oc = (open_close or "").strip()
    if oc:
        return mag if oc.startswith("开") else -mag
    d = (direction or "").strip().upper()
    if d.startswith("S") or d.startswith("卖") or "SELL" in d:
        return -mag
    return mag


def compute_position_cost_from_trades(trades: pd.DataFrame) -> pd.DataFrame:
    """从成交流水按移动加权法推算每个 ticker 当前持仓的本地币成本。

    方法（移动加权平均成本）：
        - 按 trade_date 升序处理同一 ticker 的成交。
        - 开仓/买入：加仓，累计 (数量, 成本=数量×价格)。
        - 平仓/卖出：减仓，按当前均价等比扣减成本（均价不变）。
        - 末态：cost_price_local = 剩余成本 / 剩余数量。

    返回列：ticker / net_qty / cost_price_local / cost_value_local / cost_ccy。
    仅返回末态净持仓 > 0 的 ticker。
    """
    cols = ["ticker", "net_qty", "cost_price_local", "cost_value_local", "cost_ccy"]
    if trades is None or trades.empty:
        return pd.DataFrame(columns=cols)

    out: list[dict[str, Any]] = []
    for ticker, sub in trades.groupby("ticker"):
        sub = sub.sort_values("trade_date", kind="stable")
        qty = 0.0
        cost = 0.0
        ccy: str | None = None
        for _, r in sub.iterrows():
            q = to_float(r.get("quantity")) or 0.0
            px = to_float(r.get("price_local"))
            ccy = ccy or (str(r.get("currency")).strip() if r.get("currency") else None)
            signed = _signed_qty(str(r.get("direction") or ""), str(r.get("open_close") or ""), q)
            if signed > 0:  # 加仓
                if px is not None:
                    cost += signed * px
                qty += signed
            else:  # 减仓
                dec = -signed
                if qty > 1e-9:
                    avg = cost / qty
                    take = min(dec, qty)
                    cost -= take * avg
                    qty -= take
        if qty > 1e-9:
            out.append(
                {
                    "ticker": ticker,
                    "net_qty": qty,
                    "cost_price_local": (cost / qty) if qty > 1e-9 else None,
                    "cost_value_local": cost,
                    "cost_ccy": ccy,
                }
            )
    return pd.DataFrame(out, columns=cols)
