"""每日交易情况挖掘报表。

用法:
    python scripts/daily_trade_report.py [--product 产品名|代码] [--days N]
        [--date-from YYYY-MM-DD] [--date-to YYYY-MM-DD] [--out-dir DIR] [--no-export]

产物:
    <out-dir>/daily_trades.csv          每产品 x 每日 成交汇总
    <out-dir>/monthly_trades.csv        每产品 x 每月 成交汇总
    <out-dir>/ticker_turnover.csv       每产品 x 每标的 全期成交统计

数据来源:
    - fund_transactions（来自中信 Statement / 中金"当日交易"解析入库）

注意:
    - quantity/amount_local 为交易币种口径，USD/HKD 不折算；
    - 已实现收益 realized_pnl 为交易币种口径。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd

from fundadmin.clients.compute.daily_trades import (
    daily_product_summary,
    monthly_product_summary,
    product_overview,
    ticker_turnover,
)
from fundadmin.clients.config import NAME_TO_PRODCODE, PRODCODE_TO_NAME, resolve_outputs_dir
from fundadmin.clients.store import load_transactions

pd.set_option("display.width", 200)
pd.set_option("display.max_columns", 50)


def _fmt_money(v: float) -> str:
    if pd.isna(v):
        return ""
    if abs(v) >= 1e8:
        return f"{v / 1e8:,.2f}亿"
    if abs(v) >= 1e4:
        return f"{v / 1e4:,.1f}万"
    return f"{v:,.0f}"


def _resolve_product(product: str | None) -> str | None:
    """产品名或代码 -> 产品代码。"""
    if not product:
        return None
    if product in PRODCODE_TO_NAME:
        return product
    if product in NAME_TO_PRODCODE:
        return NAME_TO_PRODCODE[product]
    raise SystemExit(f"未知产品: {product}（可选: {', '.join(PRODCODE_TO_NAME.values())}）")


def main() -> int:
    parser = argparse.ArgumentParser(description="汇总各产品每日交易情况")
    parser.add_argument("--product", default=None, help="产品名（如 铂金8号）或代码（如 SXQ602）；缺省全部")
    parser.add_argument("--days", type=int, default=15, help="控制台展示最近 N 个成交日（默认 15）")
    parser.add_argument("--date-from", default=None, help="起始日期 YYYY-MM-DD")
    parser.add_argument("--date-to", default=None, help="截止日期 YYYY-MM-DD")
    parser.add_argument("--out-dir", default=None, help="CSV 输出目录；默认 FUND_OUTPUTS_DIR/exports 或 outputs/exports")
    parser.add_argument("--no-export", action="store_true", help="只在控制台展示，不写 CSV")
    args = parser.parse_args()

    code = _resolve_product(args.product)
    label = f"{PRODCODE_TO_NAME[code]}({code})" if code else "全部产品"

    print(f"[1/4] 读取 fund_transactions [{label}] ...")
    tx = load_transactions(
        product_code=code, date_from=args.date_from, date_to=args.date_to
    )
    if tx.empty:
        print("      无交易数据")
        return 0
    print(f"      {len(tx)} 笔，覆盖 "
          f"{tx['trade_date'].nunique()} 个成交日 x {tx['product_code'].nunique()} 只产品")

    print("[2/4] 汇总 ...")
    daily = daily_product_summary(tx)
    monthly = monthly_product_summary(tx)
    tickers = ticker_turnover(tx)
    overview = product_overview(tx)

    # ---- 控制台: 产品总览 ----
    print(f"\n=== 产品全期总览 ※金额为交易币种，未折算 ===")
    show = overview.copy()
    show["first_trade"] = show["first_trade"].astype(str)
    show["last_trade"] = show["last_trade"].astype(str)
    for col in ("buy_amount", "sell_amount", "turnover", "fee", "realized_pnl"):
        show[col] = show[col].map(_fmt_money)
    show = show.rename(
        columns={
            "product_code": "代码", "product_name": "产品", "n_trades": "笔数",
            "trade_days": "成交天数", "n_tickers": "标的数", "first_trade": "首笔",
            "last_trade": "末笔", "buy_amount": "买入额", "sell_amount": "卖出额",
            "turnover": "双边成交额", "fee": "费用", "realized_pnl": "已实现收益",
        }
    )
    print(show.to_string(index=False))

    # ---- 控制台: 最近 N 个成交日 ----
    recent_dates = sorted(daily["trade_date"].unique())[-args.days:]
    recent = daily[daily["trade_date"].isin(recent_dates)]
    print(f"\n=== 最近 {len(recent_dates)} 个成交日明细 ===")
    show = recent.copy()
    show["trade_date"] = show["trade_date"].dt.date
    show = show.drop(columns=["product_code"])
    for col in ("buy_amount", "sell_amount", "net_amount", "fee", "realized_pnl"):
        show[col] = show[col].map(_fmt_money)
    show = show.rename(
        columns={
            "trade_date": "日期", "product_name": "产品", "n_trades": "笔数",
            "buy_count": "买", "sell_count": "卖", "buy_qty": "买量",
            "sell_qty": "卖量", "net_qty": "净量", "buy_amount": "买入额",
            "sell_amount": "卖出额", "net_amount": "净现金流", "fee": "费用",
            "realized_pnl": "已实现收益", "tickers": "标的",
        }
    )
    print(show.to_string(index=False))

    # ---- 控制台: 当月汇总 ----
    last_month = monthly["month"].iloc[-1]
    month_show = monthly[monthly["month"] == last_month]
    print(f"\n=== {last_month} 月度汇总 ===")
    show = month_show.copy().drop(columns=["product_code"])
    for col in ("buy_amount", "sell_amount", "net_amount", "fee", "realized_pnl"):
        show[col] = show[col].map(_fmt_money)
    show = show.rename(
        columns={
            "product_name": "产品", "trade_days": "成交天数", "n_trades": "笔数",
            "buy_amount": "买入额", "sell_amount": "卖出额", "net_amount": "净现金流",
            "fee": "费用", "realized_pnl": "已实现收益",
        }
    )
    print(show.to_string(index=False))

    # ---- 控制台: 活跃标的 TOP10 ----
    print("\n=== 活跃标的 TOP10（按卖出额）===")
    show = tickers.groupby("ticker", as_index=False).agg(
        产品数=("product_code", "nunique"),
        笔数=("n_trades", "sum"),
        卖出额=("sell_amount", "sum"),
        已实现收益=("realized_pnl", "sum"),
        标的名称=("instrument_name", "first"),
    ).sort_values("卖出额", ascending=False).head(10)
    for col in ("卖出额", "已实现收益"):
        show[col] = show[col].map(_fmt_money)
    print(show.to_string(index=False))

    if args.no_export:
        print("\n(--no-export 已跳过 CSV 输出)")
        return 0

    print("[3/4] 导出 CSV ...")
    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else resolve_outputs_dir()
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, df in (
        ("daily_trades.csv", daily),
        ("monthly_trades.csv", monthly),
        ("ticker_turnover.csv", tickers),
    ):
        path = out_dir / name
        df.to_csv(path, index=False)
        print(f"      {path}  ({len(df)} 行)")
    print("[4/4] 完成")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
