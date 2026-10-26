"""沐泽1号（SQJ420）估值表摸排报表。

沐泽1号只做 A 股 + 沪港通，持仓直接落在中金四级科目估值表中，没有逐笔
成交流水。因此"每日交易"由相邻估值日的持股数量差反推（净买入/净卖出），
并非券商成交流水口径。

用法:
    python scripts/muze_report.py [--inbox DIR] [--out-dir DIR]
        [--days N] [--no-export]

产物:
    <out-dir>/muze_nav_daily.csv        每估值日净值/规模/仓位摘要
    <out-dir>/muze_holdings.csv         每估值日每只股票持仓
    <out-dir>/muze_holdings_changes.csv 相邻估值日持仓变动（交易推断）
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd

from fundadmin.clients.config import resolve_outputs_dir
from fundadmin.portfolio.parsers.muze_valuation import (
    PRODUCT_CODE,
    PRODUCT_NAME,
    daily_trade_flow,
    holdings_changes,
    parse_muze_dir,
)

pd.set_option("display.width", 220)
pd.set_option("display.max_columns", 40)


def _wan(v: float, digits: int = 1) -> str:
    if pd.isna(v):
        return ""
    return f"{v / 1e4:,.{digits}f}万"


def _pct(v) -> str:
    return "" if pd.isna(v) else f"{v:.2f}%"


def main() -> int:
    parser = argparse.ArgumentParser(description="沐泽1号估值表摸排")
    parser.add_argument("--inbox", default="outputs/excels/fund_inbox", help="估值表根目录")
    parser.add_argument("--out-dir", default=None, help="CSV 输出目录；默认 outputs/exports")
    parser.add_argument("--days", type=int, default=10, help="控制台展示最近 N 个估值日")
    parser.add_argument("--no-export", action="store_true", help="只打印，不写 CSV")
    args = parser.parse_args()

    print(f"[1/4] 解析 {PRODUCT_NAME}({PRODUCT_CODE}) 估值表 ...")
    holdings, nav = parse_muze_dir(args.inbox)
    if nav.empty:
        print("      未找到估值表")
        return 1
    print(f"      {nav['as_of_date'].nunique()} 个估值日, {len(holdings)} 条持仓记录, "
          f"{holdings['ticker'].nunique()} 只标的")
    print(f"      区间 {nav['as_of_date'].min()} ~ {nav['as_of_date'].max()}")

    # ---- 区间概况 ----
    first, last = nav.iloc[0], nav.iloc[-1]
    print("\n=== 区间概况 ===")
    print(f"  单位净值 : {first['unit_nav']:.4f} -> {last['unit_nav']:.4f} "
          f"({last['unit_nav'] / first['unit_nav'] - 1:+.2%})")
    print(f"  累计净值 : {first['cum_nav']:.4f} -> {last['cum_nav']:.4f} "
          f"({last['cum_nav'] / first['cum_nav'] - 1:+.2%})")
    print(f"  资产净值 : {_wan(first['fund_asset_nav'])} -> {_wan(last['fund_asset_nav'])}")
    print(f"  股票市值 : {_wan(first['stock_mv'], 0)} ({first['stock_ratio_pct']:.1f}%) -> "
          f"{_wan(last['stock_mv'], 0)} ({last['stock_ratio_pct']:.1f}%)")
    print(f"  现金类   : {_wan(first['cash_like'], 0)} ({first['cash_like_ratio_pct']:.1f}%) -> "
          f"{_wan(last['cash_like'], 0)} ({last['cash_like_ratio_pct']:.1f}%)"
          "   [银行存款+券商保证金+逆回购]")
    print(f"  持股只数 : {int(first['n_stocks'])} -> {int(last['n_stocks'])}")

    # ---- 仓位结构（月内均值，避免单日资金摆布干扰） ----
    nav["month"] = nav["as_of_date"].str[:7]
    sample = nav.groupby("month").agg(
        起=("as_of_date", "first"), 止=("as_of_date", "last"),
        净值=("unit_nav", "last"), 规模=("fund_asset_nav", "last"),
        股票占比=("stock_ratio_pct", "mean"),
        保证金占比=("margin_ratio_pct", "mean"),
        逆回购占比=("repo_ratio_pct", "mean"),
        只数=("n_stocks", "last"),
    ).reset_index()
    print("\n=== 月度仓位（占比为月内均值）===")
    show = sample.copy()
    show["规模"] = show["规模"].map(_wan)
    for c in ("股票占比", "保证金占比", "逆回购占比"):
        show[c] = show[c].map(_pct)
    show["净值"] = show["净值"].map(lambda v: f"{v:.4f}")
    show.columns = ["月份", "起始日", "截止日", "单位净值", "资产净值",
                    "股票占比", "保证金占比", "逆回购占比", "持股只数"]
    print(show.to_string(index=False))

    # ---- 交易推断 ----
    chg = holdings_changes(holdings)
    print(f"\n=== 持仓变动（交易推断, 共 {len(chg)} 条, 相邻估值日净变动）===")
    recent_dates = sorted(chg["as_of_date"].unique())[-args.days:]
    show = chg[chg["as_of_date"].isin(recent_dates)].copy()
    show["chg_amount"] = show["chg_amount"].map(lambda v: f"{v / 1e4:+,.1f}万")
    show["market_value"] = show["market_value"].map(_wan)
    show["prev_date"] = show["prev_date"].fillna("-")
    show = show.rename(
        columns={
            "as_of_date": "估值日", "prev_date": "上一日", "ticker": "代码",
            "instrument_name": "名称", "market": "市场", "quantity": "持股",
            "prev_qty": "上期持股", "qty_chg": "变动股数", "price": "价格",
            "chg_amount": "变动金额", "market_value": "市值", "action": "动作",
        }
    )
    print(show.to_string(index=False))
    print("  注: 估值表非逐日披露，'上一日'跨多天时变动金额为区间净额，非单日成交。")

    # ---- 每日交易流水（买入/卖出/净额） ----
    flow = daily_trade_flow(chg)
    print(f"\n=== 每日交易推断流水（最近 {args.days} 个变动日）===")
    show = flow.tail(args.days).copy()
    for c in ("buy_amount", "sell_amount", "net_amount"):
        show[c] = show[c].map(lambda v: f"{v / 1e4:+,.1f}万")
    show = show.rename(
        columns={
            "as_of_date": "估值日", "buy_amount": "买入额", "sell_amount": "卖出额",
            "net_amount": "净买入", "n_buy": "买入笔", "n_sell": "卖出笔",
            "n_tickers": "标的数", "tickers": "涉及标的",
        }
    )
    print(show.to_string(index=False))

    # ---- 最新持仓 ----
    latest_date = holdings["as_of_date"].max()
    latest = holdings[holdings["as_of_date"] == latest_date].copy()
    latest = latest.sort_values("market_value", ascending=False)
    print(f"\n=== 最新持仓 {latest_date} ===")
    show = latest[["ticker", "instrument_name", "market", "quantity", "price",
                   "cost", "market_value", "unrealized_pnl"]].copy()
    for col in ("cost", "market_value", "unrealized_pnl"):
        show[col] = show[col].map(lambda v: f"{v / 1e4:,.1f}万")
    show.columns = ["代码", "名称", "市场", "持股", "市价", "成本", "市值", "浮动盈亏"]
    print(show.to_string(index=False))
    total = latest["market_value"].sum()
    print(f"  合计市值 {_wan(total)} | 浮动盈亏 "
          f"{_wan(latest['unrealized_pnl'].sum())}")

    # ---- 资金在各科目间的摆布（保证金 vs 逆回购） ----
    print("\n=== 资金分布（保证金 vs 逆回购, 仅列变动较大的日子）===")
    nav["cash_bucket_chg"] = nav["margin"].diff().abs() + nav["repo"].diff().abs()
    swings = nav.nlargest(8, "cash_bucket_chg").sort_values("as_of_date")
    show = swings[["as_of_date", "margin", "repo", "cash"]].copy()
    for c in ("margin", "repo", "cash"):
        show[c] = show[c].map(lambda v: f"{v / 1e4:,.0f}万")
    show.columns = ["估值日", "券商保证金", "逆回购", "银行存款"]
    print(show.to_string(index=False))

    if args.no_export:
        print("\n(--no-export 已跳过 CSV 输出)")
        return 0

    print("[2/4] 导出 CSV ...")
    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else resolve_outputs_dir()
    out_dir.mkdir(parents=True, exist_ok=True)
    nav_out = nav.drop(columns=["month"], errors="ignore")
    for name, df in (
        ("muze_nav_daily.csv", nav_out),
        ("muze_holdings.csv", holdings),
        ("muze_holdings_changes.csv", chg),
        ("muze_daily_flow.csv", flow),
    ):
        path = out_dir / name
        df.to_csv(path, index=False)
        print(f"      {path}  ({len(df)} 行)")
    print("[3/4] 完成")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
