"""产品持仓子页（分券商来源视图）。

用途:
- 选择产品 + 估值日，读每日同步落地的 fund_positions，按标的把各券商持仓股数
  透视成列（如 中金 / 中信），并展示合计股数、合计市值、权重；跨券商持有的标的
  （同一只股票同时在多个券商）高亮标出。
- 数据来源是 `fundadmin portfolio sync-latest` 自动构建的 fund_positions，
  而非手动模板上传。手动上传（写 fund_portfolio_holdings）收进底部"旧版"折叠区，
  作为不在自动同步范围内产品的兜底入口。

调用方:
- app/streamlit_app.py 的"产品持仓"tab。
"""

from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

from ...portfolio.parsers.common import clean_ticker
from ..config import PRODCODE_TO_NAME, broker_label, resolve_inbox_dir
from ..store import (
    list_position_dates,
    list_position_products,
    load_positions,
)

# 券商列展示顺序（其余券商按市值降序排在后面）。
_BROKER_ORDER = ["中金", "中信", "申万宏源", "未知"]


def _safe_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _build_by_broker_view(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """把 fund_positions 明细透视成"每标的一行、各券商股数成列"的组合持仓视图。

    返回 (展示用 DataFrame, 券商股数列名列表)。
    """
    df = df.copy()
    df["ticker"] = df["ticker"].fillna("").astype(str).str.strip()
    df["instrument_name"] = df["instrument_name"].fillna("").astype(str).str.strip()
    # 归一化代码，防御性合并：即便落库的是券商原始代码（AMZN.US vs AMZN），
    # 同一标的也能跨券商归并到一行。
    df["ticker_norm"] = df["ticker"].map(clean_ticker)
    # 标的归并键：优先归一化 ticker，缺失时用名称兜底。
    df["stock_key"] = df["ticker_norm"].where(df["ticker_norm"] != "", df["instrument_name"])
    df["broker_cn"] = df["broker"].map(broker_label)
    df["quantity"] = _safe_num(df["quantity"]).fillna(0.0)
    df["market_value_cny"] = _safe_num(df["market_value_cny"])
    df["weight"] = _safe_num(df["weight"])

    # 每标的的名称/代码/合计市值/权重。
    meta = (
        df.groupby("stock_key", as_index=False)
        .agg(
            名称=("instrument_name", "first"),
            代码=("ticker_norm", "first"),
            合计市值=("market_value_cny", "sum"),
            权重=("weight", "sum"),
        )
    )

    # 各券商股数透视：index=标的，columns=券商中文名。
    shares = df.pivot_table(
        index="stock_key",
        columns="broker_cn",
        values="quantity",
        aggfunc="sum",
    ).fillna(0.0)
    # 券商列排序：预设顺序优先，其余按列合计降序。
    present = list(shares.columns)
    ordered = [b for b in _BROKER_ORDER if b in present]
    rest = sorted([b for b in present if b not in ordered], key=lambda c: -float(shares[c].sum()))
    broker_cols = ordered + rest
    shares = shares[broker_cols]

    n_brokers = (shares != 0).sum(axis=1)
    total_shares = shares.sum(axis=1)
    shares = shares.reset_index()
    shares["合计股数"] = total_shares.values
    shares["来源券商数"] = n_brokers.values

    view = meta.merge(shares, on="stock_key").drop(columns=["stock_key"])
    view = view.sort_values("合计市值", ascending=False, na_position="last").reset_index(drop=True)

    # 列顺序：名称, 代码, <各券商股数>, 合计股数, 合计市值, 权重, 来源券商数。
    cols = ["名称", "代码", *broker_cols, "合计股数", "合计市值", "权重", "来源券商数"]
    view = view[cols]
    return view, broker_cols


def _format_view(view: pd.DataFrame, broker_cols: list[str]) -> pd.DataFrame:
    out = view.copy()
    for col in [*broker_cols, "合计股数"]:
        out[col] = _safe_num(out[col]).fillna(0).round().astype("int64").map(lambda v: f"{v:,}")
    out["合计市值"] = _safe_num(out["合计市值"]).map(lambda v: f"{v:,.2f}" if pd.notna(v) else "-")
    out["权重"] = _safe_num(out["权重"]).map(lambda v: f"{v*100:.2f}%" if pd.notna(v) else "-")
    return out


def _render_legacy_uploader() -> None:
    """旧版手动上传入口（写 fund_portfolio_holdings，不参与上方分券商视图）。"""
    st.markdown(
        "下方为**旧版手动上传**通道，写入 `fund_portfolio_holdings`（与上方自动同步的分券商视图相互独立），"
        "供未纳入自动同步的产品兜底使用。"
    )
    st.caption("列：as_of_date / product_code / broker / asset_class / ticker / instrument_name / market_value / weight / quantity")
    upload = st.file_uploader("选择持仓 Excel", type=["xlsx"], key="portfolio_uploader")
    default_as_of = st.date_input("如果文件未含 as_of_date，使用该日期", value=date.today())
    if upload is not None and st.button("写入 SQLite", key="portfolio_save"):
        inbox = resolve_inbox_dir()
        inbox.mkdir(parents=True, exist_ok=True)
        target = inbox / upload.name
        target.write_bytes(upload.getvalue())
        from ..ingest.portfolio_excel import import_portfolio_xlsx

        try:
            rows = import_portfolio_xlsx(target, as_of_override=default_as_of.isoformat())
            st.success(f"导入完成，写入 {rows} 行；源文件保存在 {target}")
        except Exception as exc:
            st.error(f"导入失败：{exc}")


def render() -> None:
    st.subheader("产品持仓 · 分券商来源")

    products = list_position_products()
    if not products:
        st.info("fund_positions 暂无数据。运行 `fundadmin portfolio sync-latest` 同步券商持仓后再查看。")
        with st.expander("旧版：手动上传持仓模板 Excel", expanded=False):
            _render_legacy_uploader()
        return

    cols = st.columns([2, 2, 3])
    with cols[0]:
        product_code = st.selectbox(
            "产品",
            options=products,
            format_func=lambda code: f"{code} {PRODCODE_TO_NAME.get(code, '')}".strip(),
        )
    dates = list_position_dates(product_code)
    with cols[1]:
        as_of = st.selectbox("估值日", options=dates) if dates else None

    df = load_positions(product_code=product_code, as_of_date=as_of) if as_of else pd.DataFrame()

    if df.empty:
        st.warning("该产品该估值日暂无持仓数据。")
        with st.expander("旧版：手动上传持仓模板 Excel", expanded=False):
            _render_legacy_uploader()
        return

    view, broker_cols = _build_by_broker_view(df)

    total_mv = _safe_num(df["market_value_cny"]).sum(skipna=True)
    n_stocks = len(view)
    n_cross = int((view["来源券商数"] > 1).sum())

    kpi = st.columns(4)
    kpi[0].metric("合计市值(CNY)", f"{total_mv:,.0f}" if pd.notna(total_mv) else "-")
    kpi[1].metric("持仓只数", str(n_stocks))
    kpi[2].metric("跨券商标的", str(n_cross))
    kpi[3].metric("券商数", str(len(broker_cols)))

    # 各券商市值占比小结。
    by_broker_mv = (
        df.assign(broker_cn=df["broker"].map(broker_label), mv=_safe_num(df["market_value_cny"]))
        .groupby("broker_cn", as_index=False)["mv"].sum()
        .sort_values("mv", ascending=False)
    )
    chips = "　".join(
        f"{r.broker_cn}: {r.mv:,.0f}" for r in by_broker_mv.itertuples() if pd.notna(r.mv)
    )
    if chips:
        st.caption(f"各券商市值(CNY)：{chips}")

    # 前 20 大（按合计权重）柱状图。
    top = view.dropna(subset=["权重"]).head(20)
    if not top.empty:
        chart_df = top.set_index(top["名称"].where(top["名称"].astype(str).str.strip() != "", top["代码"]))
        st.bar_chart(chart_df["权重"])

    st.markdown("**组合持仓（按标的，分券商股数）** — 黄色行为跨券商持有的同一标的")
    disp = _format_view(view, broker_cols)

    def _highlight_cross(row: pd.Series) -> list[str]:
        cross = str(row.get("来源券商数", "0")) not in ("0", "1")
        return ["background-color: #fff3cd" if cross else "" for _ in row]

    styler = disp.style.apply(_highlight_cross, axis=1)
    st.dataframe(styler, use_container_width=True, hide_index=True)

    with st.expander("分券商明细行（fund_positions 原始行）", expanded=False):
        detail = df.copy()
        detail["券商"] = detail["broker"].map(broker_label)
        detail_cols = {
            "券商": "券商",
            "ticker": "代码",
            "instrument_name": "名称",
            "quantity": "股数",
            "market_value_cny": "市值(CNY)",
            "weight": "权重",
            "cost_price_local": "每股成本(本币)",
            "cost_ccy": "成本币种",
            "source_files": "来源文件",
        }
        present_cols = [c for c in detail_cols if c in detail.columns]
        st.dataframe(
            detail[present_cols].rename(columns=detail_cols),
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    with st.expander("旧版：手动上传持仓模板 Excel", expanded=False):
        _render_legacy_uploader()
