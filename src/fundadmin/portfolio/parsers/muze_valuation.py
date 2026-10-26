"""沐泽1号（SQJ420，弘运盛泰）中金四级科目估值表解析。

文件形态:
    【请管理人确认估值表】<日期>弘运盛泰沐泽1号私募证券投资基金估值表_SQJ420_*.xlsx
    sheet「估值表明细表」：行0-2 表头（估值日期/单位净值），行3 列头，其后为
    四级科目明细（1002 银行存款 / 1102 股票投资 / 1202 买入返售 / 负债 / 汇总行）。

股票持仓为 1102 下的叶子科目:
    11020101601601   上交所 A 股   601601  中国太保
    11023101002371   深交所 A 股   002371  北方华创
    11024101300179   创业板        300179  四方达
    1102C101688012   科创板        688012  中微公司
    11028101H01138   沪港通        H01138  中远海能
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

PRODUCT_CODE = "SQJ420"
PRODUCT_NAME = "沐泽1号"

_SHEET_NAME = "估值表明细表"
_FIRST_DATA_ROW = 4  # 明细起始行（行0-2 表头，行3 列头）

# 1102 下的股票成本叶子: 1102 + 市场段(0101沪/3101深/4101创业/8101沪港通/C1科创) + 6位证券代码
#   1102 0101 601601   上交所 A 股  中国太保
#   1102 3101 002371   深交所 A 股  北方华创
#   1102 4101 300179   创业板      四方达
#   1102 8101 H01138   沪港通      中远海能
#   1102 C101 688012   科创板      中微公司
_STOCK_LEAF = re.compile(r"^1102(?:0101|3101|4101|8101|C101)(?P<code>[0-9A-Z]{6,7})$")

_COLS = [
    "科目代码", "科目名称", "数量", "单位成本", "成本", "成本占净值%",
    "市价", "市值", "市值占净值%", "估值增值", "停牌信息",
]

# 只统计顶层科目，避免一二级科目重复累加
_CASH_CODES = {"1002"}
_MARGIN_CODES = {"1031"}
_REPO_CODES = {"1202"}

# 尾部汇总行（名称在首列）-> summary 字段
_KEY_MAP = {
    "基金资产净值:": "fund_asset_nav",
    "累计单位净值:": "cum_nav",
    "实收资本金额": "share_capital",
    "本年已实现收益:": "ytd_realized",
    "从成立日开始的累计已实现收益:": "cum_realized",
    "累计派现金额:": "payout",
    "期初单位净值:": "open_nav",
    "昨日单位净值:": "prev_nav",
}


def _market_of(code: str) -> str:
    if code.startswith("H"):
        return "沪港通"
    if code.startswith("688"):
        return "科创板"
    if code.startswith("300"):
        return "深交所_创业板"
    if code.startswith("60"):
        return "上交所"
    return "深交所"


def _norm_ticker(code: str) -> str:
    """H01138 -> 01138.HK；A 股保持 6 位。"""
    return f"{code[1:]}.HK" if code.startswith("H") else code


def _to_float(v) -> float:
    try:
        if v is None or pd.isna(v):
            return 0.0
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return 0.0


def _row_numbers(values: list) -> list[float]:
    """取一行里所有可转数值的单元格（跳过列头文字/千分位字符串均可）。"""
    out = []
    for v in values:
        try:
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            out.append(float(str(v).replace(",", "")))
        except (TypeError, ValueError):
            continue
    return out


def _parse_header(raw: pd.DataFrame) -> tuple[str, float | None]:
    """从顶部文字区提取 估值日期 / 单位净值。"""
    text = " ".join(str(v) for v in raw.iloc[:4].to_numpy().ravel() if str(v) != "nan")
    m_date = re.search(r"(\d{4}-\d{2}-\d{2})", text)
    m_nav = re.search(r"单位净值[：:]\s*([\d.]+)", text)
    return (m_date.group(1) if m_date else ""), (float(m_nav.group(1)) if m_nav else None)


def parse_muze_valuation(path: str | Path) -> tuple[pd.DataFrame, dict]:
    """解析单份沐泽1号估值表，返回 (holdings, summary)。

    holdings: 每只股票一行: as_of_date/ticker/account_code/instrument_name/
        market/quantity/cost/price/market_value/unrealized_pnl
    summary:  每估值日一行的资产/净值/损益指标。
    """
    path = Path(path)
    raw = pd.read_excel(path, sheet_name=_SHEET_NAME, header=None)
    as_of_date, unit_nav = _parse_header(raw)

    body = raw.iloc[_FIRST_DATA_ROW:].copy()
    body.columns = _COLS

    rows = []
    cash = repo = margin = 0.0
    summary: dict[str, object] = {
        "as_of_date": as_of_date,
        "product_code": PRODUCT_CODE,
        "product_name": PRODUCT_NAME,
        "unit_nav": unit_nav,
        "cash": 0.0,
        "repo": 0.0,
        "margin": 0.0,
        "source_file": path.name,
    }

    for r in body.itertuples(index=False):
        code = str(r[0]).strip()
        name = str(r[1]).strip()

        leaf = _STOCK_LEAF.match(code)
        if leaf:
            rows.append(
                {
                    "as_of_date": as_of_date,
                    "account_code": leaf.group("code"),
                    "ticker": _norm_ticker(leaf.group("code")),
                    "instrument_name": name,
                    "market": _market_of(leaf.group("code")),
                    "quantity": _to_float(r[2]),
                    "cost": _to_float(r[4]),
                    "price": _to_float(r[6]),
                    "market_value": _to_float(r[7]),
                    "unrealized_pnl": _to_float(r[9]),
                }
            )
            continue

        nums = _row_numbers([r[i] for i in range(2, len(_COLS))])
        # 顶层科目只取一行，避免一二级科目重复累加
        if code in _CASH_CODES:
            cash += max(nums) if nums else 0.0
        elif code in _MARGIN_CODES:
            margin += max(nums) if nums else 0.0
        elif code in _REPO_CODES:
            repo += max(nums) if nums else 0.0
        elif code or name:
            # 尾部汇总行: 名称在首列，数值散落在其后（文本数字带千分位）
            nums = _row_numbers([r[i] for i in range(1, len(_COLS))])
            if code in _KEY_MAP and nums:
                # 基金资产净值取市值口径（科目现值最大的那个数）
                summary[_KEY_MAP[code]] = max(nums) if code == "基金资产净值:" else nums[0]

    holdings = pd.DataFrame(rows)
    if not holdings.empty:
        holdings["product_code"] = PRODUCT_CODE
        holdings["product_name"] = PRODUCT_NAME
        holdings["source_file"] = path.name

    summary["cash"] = cash
    summary["repo"] = repo
    summary["margin"] = margin

    if not holdings.empty:
        nav = summary.get("fund_asset_nav") or 0.0
        cash_like = cash + margin + repo
        summary["stock_mv"] = float(holdings["market_value"].sum())
        summary["stock_cost"] = float(holdings["cost"].sum())
        summary["stock_unrealized"] = float(holdings["unrealized_pnl"].sum())
        summary["n_stocks"] = int(len(holdings))
        summary["cash_like"] = cash_like
        summary["cash_ratio_pct"] = 100.0 * cash / nav if nav else None
        summary["margin_ratio_pct"] = 100.0 * margin / nav if nav else None
        summary["repo_ratio_pct"] = 100.0 * repo / nav if nav else None
        summary["cash_like_ratio_pct"] = 100.0 * cash_like / nav if nav else None
        summary["stock_ratio_pct"] = 100.0 * summary["stock_mv"] / nav if nav else None

    return holdings, summary


def parse_muze_dir(inbox_dir: str | Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """解析 fund_inbox 下全部沐泽1号估值表（按估值日去重，优先非 _staged）。

    返回: (holdings, nav_summary)
    """
    by_date: dict[str, Path] = {}
    for f in sorted(Path(inbox_dir).glob("**/*SQJ420*.xlsx")):
        m = re.search(r"_(\d{8})\.xlsx$", f.name)
        if not m:
            continue
        d = f"{m.group(1)[:4]}-{m.group(1)[4:6]}-{m.group(1)[6:]}"
        prev = by_date.get(d)
        if prev is None or ("_staged" in str(prev) and "_staged" not in str(f)):
            by_date[d] = f

    holdings, summaries = [], []
    for d, f in sorted(by_date.items()):
        h, s = parse_muze_valuation(f)
        if not h.empty:
            holdings.append(h)
        summaries.append(s)

    hdf = pd.concat(holdings, ignore_index=True) if holdings else pd.DataFrame()
    sdf = pd.DataFrame(summaries).sort_values("as_of_date").reset_index(drop=True)
    return hdf, sdf


def holdings_changes(holdings: pd.DataFrame) -> pd.DataFrame:
    """相邻估值日持仓变动 -> 净买入/卖出、新进/清仓、市值变动。

    以「全部估值日 x 全部标的」补齐为面板，缺失持仓视为 0，
    避免「清仓后隔日再买回」被误判为一次减仓。
    qty_chg = 当日数量 - 上一估值日数量；新进个股按当日数量全额计入。
    chg_amount = qty_chg * 参考价（清仓用最后已知价，近似，不含费用）。
    首日存量标记为「期初」，不计为交易。
    """
    if holdings.empty:
        return pd.DataFrame()

    df = holdings.copy()
    dates = sorted(df["as_of_date"].unique())
    tickers = sorted(df["ticker"].unique())
    idx = pd.MultiIndex.from_product([dates, tickers], names=["as_of_date", "ticker"])
    piv = df.set_index(["as_of_date", "ticker"])

    qty = piv["quantity"].reindex(idx, fill_value=0.0)
    mv = piv["market_value"].reindex(idx, fill_value=0.0)
    price = piv["price"].reindex(idx)
    name = piv["instrument_name"].reindex(idx)
    market = piv["market"].reindex(idx)

    panel = pd.DataFrame({"quantity": qty, "market_value": mv, "price": price,
                          "instrument_name": name, "market": market}).reset_index()
    for col in ("price", "instrument_name", "market"):
        panel[col] = panel.groupby("ticker")[col].ffill().bfill()
    g = panel.groupby("ticker")
    panel["prev_qty"] = g["quantity"].shift(1)
    panel["prev_mv"] = g["market_value"].shift(1)
    panel["prev_date"] = g["as_of_date"].shift(1)

    is_first = panel["as_of_date"] == dates[0]
    panel["qty_chg"] = panel["quantity"] - panel["prev_qty"].fillna(0.0)
    panel["mv_chg"] = panel["market_value"] - panel["prev_mv"].fillna(0.0)
    panel["chg_amount"] = panel["qty_chg"] * panel["price"]

    panel["action"] = "加仓"
    panel.loc[is_first, "action"] = "期初"
    panel.loc[~is_first & (panel["quantity"] > 0) & (panel["prev_qty"] == 0), "action"] = "新进"
    panel.loc[~is_first & (panel["quantity"] == 0) & (panel["prev_qty"] > 0), "action"] = "清仓"
    panel.loc[~is_first & (panel["qty_chg"] < 0) & (panel["quantity"] > 0), "action"] = "减仓"

    moved = panel[(panel["qty_chg"].abs() > 1e-9) & (~is_first)].copy()
    moved = moved.sort_values(["as_of_date", "chg_amount"], ascending=[True, False])
    return moved[
        [
            "as_of_date", "prev_date", "ticker", "instrument_name", "market",
            "quantity", "prev_qty", "qty_chg", "price", "chg_amount",
            "market_value", "mv_chg", "action",
        ]
    ].reset_index(drop=True)


def daily_trade_flow(changes: pd.DataFrame) -> pd.DataFrame:
    """把持仓变动聚合成「每日交易情况」：买入额/卖出额/净买入/标的数。

    金额为按估值日价格估算的区间净额（估值表非逐日披露）。
    """
    if changes.empty:
        return pd.DataFrame(
            columns=["as_of_date", "buy_amount", "sell_amount", "net_amount",
                     "n_buy", "n_sell", "n_tickers", "tickers"]
        )
    df = changes.copy()
    df["buy_amount"] = df["chg_amount"].where(df["chg_amount"] > 0, 0.0)
    df["sell_amount"] = (-df["chg_amount"]).where(df["chg_amount"] < 0, 0.0)
    out = (
        df.groupby("as_of_date")
        .apply(
            lambda g: pd.Series(
                {
                    "buy_amount": g["buy_amount"].sum(),
                    "sell_amount": g["sell_amount"].sum(),
                    "net_amount": g["buy_amount"].sum() - g["sell_amount"].sum(),
                    "n_buy": int((g["chg_amount"] > 0).sum()),
                    "n_sell": int((g["chg_amount"] < 0).sum()),
                    "n_tickers": g["ticker"].nunique(),
                    "tickers": " ".join(sorted(g["ticker"].unique())),
                }
            ),
            include_groups=False,
        )
        .reset_index()
    )
    return out.sort_values("as_of_date").reset_index(drop=True)


__all__ = [
    "parse_muze_valuation",
    "parse_muze_dir",
    "holdings_changes",
    "daily_trade_flow",
    "PRODUCT_CODE",
    "PRODUCT_NAME",
]
