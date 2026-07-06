"""中信（CITIC）收益互换估值报告（协会版本）解析器。

中国证券业协会标准多 sheet 格式，中信证券自 2026-06 起以"试运行"邮件下发，
逐步替代旧的逐币种"履约保障报告（Statement）"。文件名形如：

    试运行-【中信证券】【107244】【...】-收益互换估值报告-【估值日2026-06-25】【协会版本】_..xlsx

sheet 结构：
    客户估值汇总 / 存续合约明细 / 存续标的汇总 / 公司行为 / 资金变动明细

底层标的持仓在「存续标的汇总」sheet（表头在第 2 行）。同一标的因不同合约批次拆成多行，
需按标的代码汇总。关键列：

    标的代码 / 标的名称 / 标的方向（选填）/ 名义数量 / 计价货币 / 参考汇率 /
    标的名义本金 / 估值价格 / 标的市值

市值换算（已与同日 USD 履约保障 Statement 的标的市值逐一核对一致）：
    市值（人民币）= 标的市值（计价货币）× 参考汇率
其中「参考汇率」为 交易货币 → 人民币（USD≈6.80、HKD≈0.867）。

与 Statement 的关系：本表为单文件覆盖该互换全部币种（USD+HKD）的"超集"，旧 USD
Statement 仅是其 USD 子集。故协会版本存在时应作为中信唯一持仓来源，抑制 Statement，
避免重复计算（详见 cross_broker_report.build_cross_broker_report）。
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from fundadmin.portfolio.parsers.common import clean_ticker, to_float, to_int

HOLDINGS_SHEET = "存续标的汇总"
SUMMARY_SHEET = "客户估值汇总"

# 「存续标的汇总」表头别名（去空格后比较）
CODE_LABELS = {"标的代码"}
NAME_LABELS = {"标的名称"}
QTY_LABELS = {"名义数量"}
CCY_LABELS = {"计价货币"}
FX_LABELS = {"参考汇率"}
MV_LABELS = {"标的市值"}
NOTIONAL_LABELS = {"标的名义本金"}
DIRECTION_LABELS = {"标的方向（选填）", "标的方向(选填)", "标的方向"}

_OUT_COLS = [
    "ticker", "company", "shares", "market_value_cny",
    "cost_price_local", "cost_value_local", "cost_ccy", "source_file",
]


def _norm(text: object) -> str:
    return str(text or "").strip().replace(" ", "").replace("\n", "")


def is_citic_assoc_report(path: Path) -> bool:
    """按 sheet 结构判断是否为中信收益互换协会版本估值报告（不依赖文件名）。"""
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False
    try:
        names = {_norm(s) for s in pd.ExcelFile(path).sheet_names}
    except Exception:
        return False
    return SUMMARY_SHEET in names and HOLDINGS_SHEET in names


def _find_sheet(path: Path, target: str) -> str | int:
    for name in pd.ExcelFile(path).sheet_names:
        if _norm(name) == target:
            return name
    raise ValueError(f"协会版本估值表缺少「{target}」sheet: {path}")


def _detect_header_row(df: pd.DataFrame) -> int:
    """定位「存续标的汇总」表头行（含标的代码 + 标的名称）。"""
    for idx in range(min(10, len(df))):
        texts = {_norm(v) for v in df.iloc[idx].tolist()}
        if (texts & CODE_LABELS) and (texts & NAME_LABELS):
            return idx
    raise ValueError("协会版本「存续标的汇总」未找到表头行（标的代码/标的名称）")


def parse_citic_assoc_holdings(path: Path) -> pd.DataFrame:
    """解析协会版本「存续标的汇总」，按标的代码汇总为标准化持仓 DataFrame。

    返回列：
        ticker / company / shares / market_value_cny /
        cost_price_local / cost_value_local / cost_ccy / source_file

    说明：
        - 同一标的的多个合约批次按标的代码合并（名义数量、标的市值、标的名义本金求和）。
        - 「标的方向」为"空"的行，名义数量与市值取负（净敞口口径）；现有数据均为多头。
        - 已平仓占位行（名义数量与标的市值均为 0）在合并后被过滤。
        - cost 为交易货币口径：cost_value_local=标的名义本金；cost_price_local=每股成本。
    """
    sheet = _find_sheet(path, HOLDINGS_SHEET)
    df_raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object, engine="openpyxl")
    if df_raw.empty:
        raise ValueError(f"协会版本「存续标的汇总」为空: {path}")

    header_row = _detect_header_row(df_raw)
    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = [_norm(v) for v in df_raw.iloc[header_row].tolist()]

    def _pick(labels: set[str]) -> str | None:
        for col in df.columns:
            if col in labels:
                return col
        return None

    code_col = _pick(CODE_LABELS)
    name_col = _pick(NAME_LABELS)
    qty_col = _pick(QTY_LABELS)
    fx_col = _pick(FX_LABELS)
    mv_col = _pick(MV_LABELS)
    if not code_col or not name_col or not qty_col or not fx_col or not mv_col:
        raise ValueError(f"协会版本「存续标的汇总」缺少必要列: {path}")
    ccy_col = _pick(CCY_LABELS)
    notional_col = _pick(NOTIONAL_LABELS)
    dir_col = _pick(DIRECTION_LABELS)

    rows: list[dict[str, object]] = []
    for _, row in df.iterrows():
        code = _norm(row.get(code_col))
        if not code:
            continue
        qty = to_float(row.get(qty_col))
        mv_local = to_float(row.get(mv_col))
        fx = to_float(row.get(fx_col))
        if qty is None or fx is None:
            continue
        # 方向：空头取负敞口（现有数据均为多头，此分支为前向兼容）。
        sign = -1.0 if (dir_col and "空" in _norm(row.get(dir_col))) else 1.0
        notional = to_float(row.get(notional_col)) if notional_col else None
        rows.append(
            {
                "ticker": clean_ticker(code),
                "company": str(row.get(name_col) or "").strip(),
                "ccy": str(row.get(ccy_col) or "").strip() if ccy_col else "",
                "shares_signed": sign * qty,
                "mv_cny": sign * (mv_local * fx) if mv_local is not None else 0.0,
                "cost_local": sign * notional if notional is not None else None,
            }
        )

    if not rows:
        return pd.DataFrame(columns=_OUT_COLS)

    raw = pd.DataFrame(rows)
    grouped = raw.groupby("ticker", as_index=False).agg(
        company=("company", "first"),
        cost_ccy=("ccy", "first"),
        shares_signed=("shares_signed", "sum"),
        mv_cny=("mv_cny", "sum"),
        cost_local=("cost_local", "sum"),
    )
    # 过滤已平仓占位（股数与市值均为 0）。
    grouped = grouped[(grouped["shares_signed"] != 0) | (grouped["mv_cny"] != 0)].copy()
    if grouped.empty:
        return pd.DataFrame(columns=_OUT_COLS)

    out = pd.DataFrame()
    out["ticker"] = grouped["ticker"]
    out["company"] = grouped["company"]
    out["shares"] = grouped["shares_signed"].map(to_int)
    out["market_value_cny"] = grouped["mv_cny"]
    out["cost_value_local"] = grouped["cost_local"]
    out["cost_price_local"] = [
        (cv / sh) if (cv is not None and not pd.isna(cv) and sh) else None
        for cv, sh in zip(grouped["cost_local"], grouped["shares_signed"])
    ]
    out["cost_ccy"] = grouped["cost_ccy"].replace("", None)
    out["source_file"] = path.name
    return out[_OUT_COLS]
