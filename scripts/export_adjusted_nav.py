"""导出复权净值曲线与分红事件明细。

用法:
    python scripts/export_adjusted_nav.py [--out-dir DIR]

产物:
    <out-dir>/<code>_<产品名>_复权净值走势.png   原始净值 vs 复权净值双线图
    <out-dir>/adjusted_nav.csv                   复权净值明细
    <out-dir>/dividend_events.csv                分红事件明细

数据来源:
    - 累计派现: raw_sheet_rows 中的券商估值表（`累计派现金额` 行）
    - 单位净值: product_valuation
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pandas as pd

from fundadmin.clients.config import resolve_outputs_dir
from fundadmin.portfolio.adjusted_nav import (
    build_nav_variants,
    detect_dividend_events,
    load_nav_history,
    load_payout_series,
)
from fundadmin.portfolio.viz import generate_adjusted_nav_chart


def main() -> int:
    parser = argparse.ArgumentParser(description="导出复权净值曲线")
    parser.add_argument(
        "--out-dir",
        default=None,
        help="输出目录；默认 FUND_OUTPUTS_DIR 或 outputs/exports",
    )
    args = parser.parse_args()

    out_dir = Path(args.out_dir).expanduser().resolve() if args.out_dir else resolve_outputs_dir()
    out_dir.mkdir(parents=True, exist_ok=True)

    print("[1/4] 读取累计派现 ...")
    payout = load_payout_series()
    print(f"      {len(payout)} 行，覆盖 {payout['product_code'].nunique()} 只产品")

    print("[2/4] 读取净值序列 ...")
    nav = load_nav_history()
    print(f"      {len(nav)} 行")

    print("[3/4] 计算复权净值与分红事件 ...")
    adjusted = build_nav_variants(nav, payout)
    events = detect_dividend_events(payout, nav_window=nav)
    in_window = events[events["in_nav_window"]]

    adj_path = out_dir / "adjusted_nav.csv"
    ev_path = out_dir / "dividend_events.csv"
    adjusted.to_csv(adj_path, index=False)
    events.to_csv(ev_path, index=False)
    print(f"      明细: {adj_path}")
    print(f"      分红: {ev_path}")

    print("[4/4] 绘制曲线 ...")
    for code, group in adjusted.groupby("product_code"):
        name = group["product_name"].iloc[0]
        path = generate_adjusted_nav_chart(
            group,
            product_code=code,
            product_name=name,
            events=events,
            out_path=out_dir / f"{code}_{name}_复权净值走势.png",
        )
        print(f"      {path}")

    print("\n=== 分红事件（落在净值窗口内）===")
    if in_window.empty:
        print("  无")
    else:
        show = in_window[
            ["product_code", "product_name", "val_date", "amount_per_unit",
             "cum_payout_after", "nav_change", "explained_by_payout"]
        ].copy()
        show["val_date"] = show["val_date"].dt.date
        if "nav_change" in show:
            show["nav_change"] = show["nav_change"].map(lambda v: f"{v:+.2%}" if pd.notna(v) else "")
        print(show.to_string(index=False))
        flags = in_window["explained_by_payout"].fillna(False).astype(bool)
        n_unchecked = int((~flags).sum())
        if n_unchecked:
            print(
                f"\n  提示: {n_unchecked} 笔事件的当日净值跌幅与分红金额不匹配"
                "（多为 `累计派现金额` 登记日与净值扣减日相差 1 个交易日），建议人工复核。"
            )

    out_of_window = events[~events["in_nav_window"]]
    if not out_of_window.empty:
        print("\n=== 分红事件（净值窗口之前，仅作复权基期平移）===")
        for row in out_of_window.itertuples():
            print(
                f"  {row.product_code} {row.product_name} {row.val_date.date()} "
                f"每份 {row.amount_per_unit:.4f}"
            )

    print("\n=== 三种口径区间收益（各自首日归一）===")
    rows = []
    for code, group in adjusted.groupby("product_code"):
        group = group.sort_values("val_date")
        name = group["product_name"].iloc[0]
        base = group.iloc[0]
        last = group.iloc[-1]

        def _r(col):
            return float(last[col] / base[col] - 1.0)

        rows.append(
            (
                code,
                name,
                f"{pd.Timestamp(base['val_date']).date()} ~ {pd.Timestamp(last['val_date']).date()}",
                _r("unit_nav"),
                _r("adj_nav"),
                _r("reinvest_nav"),
                _r("reinvest_nav") - _r("adj_nav"),
            )
        )
    table = pd.DataFrame(
        rows,
        columns=["code", "name", "区间", "未复权", "算术加回", "分红再投资", "再投资-算术"],
    )
    for col in ("未复权", "算术加回", "分红再投资", "再投资-算术"):
        table[col] = table[col].map(lambda v: f"{v:+.2%}")
    print(table.to_string(index=False))
    print(
        "\n注: 无分红产品两种复权口径应完全一致（差额 0.00pp），可作为交叉校验。\n"
        "    算术加回在基期已含历史派现时会稀释区间收益，故以分红再投资为准。"
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
