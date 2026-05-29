"""产品持仓子页。

用途:
- 选择产品 + 估值日，展示持仓表与前 N 大柱状图；提供模板 Excel 上传写入。

调用方:
- apps/streamlit/internal_dashboard_app.py 的 tab2
"""

from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

from ..config import PRODCODE_TO_NAME, resolve_inbox_dir
from ..store import list_holdings_dates, list_holdings_products, load_holdings


def _safe_format(value: object) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{float(value):,.2f}"


def render() -> None:
    st.subheader("产品持仓")

    products = list_holdings_products() or list(PRODCODE_TO_NAME.keys())
    dates = list_holdings_dates()

    cols = st.columns([2, 2, 3])
    with cols[0]:
        product_code = st.selectbox(
            "产品",
            options=products,
            format_func=lambda code: f"{code} {PRODCODE_TO_NAME.get(code, '')}".strip(),
        )
    with cols[1]:
        as_of = st.selectbox("估值日", options=dates) if dates else None

    df = load_holdings(product_code=product_code, as_of_date=as_of) if as_of else pd.DataFrame()

    kpi = st.columns(4)
    if df.empty:
        kpi[0].metric("总市值", "-")
        kpi[1].metric("股票占比", "-")
        kpi[2].metric("现金占比", "-")
        kpi[3].metric("持仓条数", "0")
    else:
        df["market_value"] = pd.to_numeric(df["market_value"], errors="coerce")
        df["weight"] = pd.to_numeric(df["weight"], errors="coerce")
        total_mv = df["market_value"].sum(skipna=True)
        equity_mask = df["asset_class"].astype(str).str.contains("股", na=False)
        cash_mask = df["asset_class"].astype(str).str.contains("现金", na=False)
        equity_w = df.loc[equity_mask, "weight"].sum(skipna=True)
        cash_w = df.loc[cash_mask, "weight"].sum(skipna=True)
        kpi[0].metric("总市值", _safe_format(total_mv))
        kpi[1].metric("股票占比", f"{equity_w*100:.2f}%" if pd.notna(equity_w) else "-")
        kpi[2].metric("现金占比", f"{cash_w*100:.2f}%" if pd.notna(cash_w) else "-")
        kpi[3].metric("持仓条数", str(len(df)))

        top = df.dropna(subset=["weight"]).sort_values("weight", ascending=False).head(20)
        if not top.empty:
            chart_df = top.set_index(top["instrument_name"].where(top["instrument_name"].astype(str).str.strip().ne(""), top["ticker"]))
            st.bar_chart(chart_df["weight"])

        rename_map = {
            "ticker": "代码",
            "instrument_name": "名称",
            "asset_class": "资产类别",
            "broker": "券商",
            "market_value": "市值",
            "weight": "权重",
            "quantity": "数量",
            "as_of_date": "估值日",
        }
        st.dataframe(df.rename(columns=rename_map), use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("**上传持仓模板 Excel**（列：as_of_date / product_code / broker / asset_class / ticker / instrument_name / market_value / weight / quantity）")
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
