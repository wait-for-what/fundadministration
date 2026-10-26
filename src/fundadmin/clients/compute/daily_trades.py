"""每日交易情况挖掘: 从 fund_transactions 生成多维度汇总。

口径说明:
- quantity/amount_local 在中信数据中带符号（卖出为负）；汇总时买入/卖出分组
  均取绝对值，净额 = 买入 - 卖出。
- fee 为佣金+印花税合计；realized_pnl 为已实现收益（交易币种）。
- 跨币种（USD/HKD）不折算，按交易币种分组或以 currencies 列标注。
"""

from __future__ import annotations

import pandas as pd

SIDE_BUY = "B"
SIDE_SELL = "S"


def _empty_tx() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "trade_date", "product_code", "product_name", "broker", "ticker",
            "instrument_name", "direction", "open_close", "price_local",
            "quantity", "amount_local", "fee", "realized_pnl", "currency",
        ]
    )


def daily_product_summary(tx: pd.DataFrame) -> pd.DataFrame:
    """每产品 x 每日 成交汇总（全账户合计）。"""
    if tx.empty:
        return pd.DataFrame(
            columns=[
                "trade_date", "product_code", "product_name", "n_trades",
                "buy_count", "sell_count",
                "buy_qty", "sell_qty", "net_qty",
                "buy_amount", "sell_amount", "net_amount",
                "fee", "realized_pnl", "currencies", "tickers",
            ]
        )

    df = tx.copy()
    df["trade_date"] = pd.to_datetime(df["trade_date"])
    is_buy = df["direction"].isin([SIDE_BUY, "买入"])
    is_sell = df["direction"].isin([SIDE_SELL, "卖出"])

    for col in ("quantity", "amount_local", "fee", "realized_pnl"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    out = (
        df.groupby(["trade_date", "product_code", "product_name"])
        .apply(
            lambda g: pd.Series(
                {
                    "n_trades": len(g),
                    "buy_count": int(is_buy[g.index].sum()),
                    "sell_count": int(is_sell[g.index].sum()),
                    "buy_qty": g.loc[is_buy.loc[g.index], "quantity"].abs().sum(),
                    "sell_qty": g.loc[is_sell.loc[g.index], "quantity"].abs().sum(),
                    "buy_amount": g.loc[is_buy.loc[g.index], "amount_local"].abs().sum(),
                    "sell_amount": g.loc[is_sell.loc[g.index], "amount_local"].abs().sum(),
                    "net_qty": g.loc[is_buy.loc[g.index], "quantity"].abs().sum()
                    - g.loc[is_sell.loc[g.index], "quantity"].abs().sum(),
                    "net_amount": g.loc[is_buy.loc[g.index], "amount_local"].abs().sum()
                    - g.loc[is_sell.loc[g.index], "amount_local"].abs().sum(),
                    "fee": g["fee"].sum(),
                    "realized_pnl": g["realized_pnl"].sum(),
                    "currencies": "/".join(sorted(g["currency"].dropna().unique())),
                    "tickers": " ".join(sorted(g["ticker"].dropna().unique())),
                }
            ),
            include_groups=False,
        )
        .reset_index()
    )
    return out.sort_values(["trade_date", "product_code"]).reset_index(drop=True)


def monthly_product_summary(tx: pd.DataFrame) -> pd.DataFrame:
    """每产品 x 每月 成交汇总。"""
    if tx.empty:
        return pd.DataFrame(
            columns=[
                "month", "product_code", "product_name", "trade_days", "n_trades",
                "buy_amount", "sell_amount", "net_amount", "fee", "realized_pnl",
            ]
        )

    df = tx.copy()
    df["trade_date"] = pd.to_datetime(df["trade_date"])
    is_buy = df["direction"].isin([SIDE_BUY, "买入"])
    df["month"] = df["trade_date"].dt.strftime("%Y-%m")
    for col in ("amount_local", "fee", "realized_pnl"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    out = (
        df.groupby(["month", "product_code", "product_name"])
        .agg(
            trade_days=("trade_date", "nunique"),
            n_trades=("ticker", "size"),
            buy_amount=("amount_local", lambda s: s[is_buy[s.index]].abs().sum()),
            sell_amount=("amount_local", lambda s: s[~is_buy[s.index]].abs().sum()),
            fee=("fee", "sum"),
            realized_pnl=("realized_pnl", "sum"),
        )
        .reset_index()
    )
    out["net_amount"] = out["buy_amount"] - out["sell_amount"]
    return out.sort_values(["month", "product_code"]).reset_index(drop=True)


def ticker_turnover(tx: pd.DataFrame) -> pd.DataFrame:
    """每产品 x 每标的 全期成交统计（用于挖掘活跃标的与累计损益）。"""
    if tx.empty:
        return pd.DataFrame(
            columns=[
                "product_code", "product_name", "ticker", "instrument_name",
                "n_trades", "buy_count", "sell_count", "buy_qty", "sell_qty",
                "net_qty", "buy_amount", "sell_amount", "fee", "realized_pnl",
                "first_trade_date", "last_trade_date",
            ]
        )

    df = tx.copy()
    df["trade_date"] = pd.to_datetime(df["trade_date"])
    is_buy = df["direction"].isin([SIDE_BUY, "买入"])
    for col in ("quantity", "amount_local", "fee", "realized_pnl"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    out = (
        df.groupby(["product_code", "product_name", "ticker", "instrument_name"])
        .apply(
            lambda g: pd.Series(
                {
                    "n_trades": len(g),
                    "buy_count": int(is_buy[g.index].sum()),
                    "sell_count": int((~is_buy[g.index]).sum()),
                    "buy_qty": g.loc[is_buy.loc[g.index], "quantity"].abs().sum(),
                    "sell_qty": g.loc[~is_buy.loc[g.index], "quantity"].abs().sum(),
                    "buy_amount": g.loc[is_buy.loc[g.index], "amount_local"].abs().sum(),
                    "sell_amount": g.loc[~is_buy.loc[g.index], "amount_local"].abs().sum(),
                    "fee": g["fee"].sum(),
                    "realized_pnl": g["realized_pnl"].sum(),
                    "net_qty": g.loc[is_buy.loc[g.index], "quantity"].abs().sum()
                    - g.loc[~is_buy.loc[g.index], "quantity"].abs().sum(),
                    "first_trade_date": g["trade_date"].min(),
                    "last_trade_date": g["trade_date"].max(),
                }
            ),
            include_groups=False,
        )
        .reset_index()
    )
    return out.sort_values(["product_code", "sell_amount"], ascending=[True, False]).reset_index(drop=True)


def product_overview(tx: pd.DataFrame) -> pd.DataFrame:
    """每产品全期总览（笔数、成交额、费用、已实现损益、活跃天数、标的数）。"""
    if tx.empty:
        return pd.DataFrame()

    df = tx.copy()
    df["trade_date"] = pd.to_datetime(df["trade_date"])
    is_buy = df["direction"].isin([SIDE_BUY, "买入"])
    for col in ("quantity", "amount_local", "fee", "realized_pnl"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    out = (
        df.groupby(["product_code", "product_name"])
        .apply(
            lambda g: pd.Series(
                {
                    "n_trades": len(g),
                    "trade_days": g["trade_date"].nunique(),
                    "n_tickers": g["ticker"].nunique(),
                    "first_trade": g["trade_date"].min().date(),
                    "last_trade": g["trade_date"].max().date(),
                    "buy_amount": g.loc[is_buy.loc[g.index], "amount_local"].abs().sum(),
                    "sell_amount": g.loc[~is_buy.loc[g.index], "amount_local"].abs().sum(),
                    "fee": g["fee"].sum(),
                    "realized_pnl": g["realized_pnl"].sum(),
                }
            ),
            include_groups=False,
        )
        .reset_index()
    )
    out["turnover"] = out["buy_amount"] + out["sell_amount"]
    return out.sort_values("turnover", ascending=False).reset_index(drop=True)


__all__ = [
    "daily_product_summary",
    "monthly_product_summary",
    "ticker_turnover",
    "product_overview",
]
