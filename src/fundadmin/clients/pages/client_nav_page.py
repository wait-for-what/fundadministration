"""客户净值子页。

用途:
- 选择估值日，展示客户 × 产品的份额 / 单位净值 / 估算市值；
  支持下载 CSV，并提供"上传新一日净值 Excel"按钮。

调用方:
- apps/streamlit/internal_dashboard_app.py 的 tab3
"""

from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

from ..compute import build_client_nav_frame
from ..config import resolve_inbox_dir
from ..store import list_nav_dates


def _format_money(v: object) -> str:
    if v is None or pd.isna(v):
        return "-"
    return f"{float(v):,.2f}"


def render() -> None:
    st.subheader("客户净值")

    nav_dates = list_nav_dates()
    chosen = st.selectbox("估值日（缺省取每个产品的最新一日）", options=["最新"] + nav_dates)
    as_of = None if chosen == "最新" else chosen

    df = build_client_nav_frame(as_of_date=as_of)
    if df.empty:
        st.info("尚无客户或净值数据，请先 import-clients 与 import-nav。")
    else:
        total_mv = pd.to_numeric(df["market_value"], errors="coerce").sum(skipna=True)
        kpi = st.columns(3)
        kpi[0].metric("客户数", df["custname"].nunique())
        kpi[1].metric("产品条目数", str(len(df)))
        kpi[2].metric("合计市值", _format_money(total_mv))

        display = df.copy()
        display["持有份额"] = display["holding_shares"].apply(lambda v: f"{v:,.2f}" if pd.notna(v) else "-")
        display["单位净值"] = display["nav_unit"].apply(lambda v: f"{v:,.4f}" if pd.notna(v) else "-")
        display["市值(估算)"] = display["market_value"].apply(_format_money)
        cols_show = ["custname", "prodcode", "prodname", "持有份额", "单位净值", "市值(估算)", "as_of_date", "email"]
        rename = {"custname": "客户", "prodcode": "产品代码", "prodname": "产品名称", "as_of_date": "估值日", "email": "邮箱"}
        st.dataframe(display[cols_show].rename(columns=rename), use_container_width=True, hide_index=True)

        st.download_button(
            "下载 CSV",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"client_nav_{as_of or 'latest'}.csv",
            mime="text/csv",
        )

    st.divider()
    st.markdown("**上传产品净值 Excel**（列：prodcode 或 product_name + nav_unit/unit_nav，可选 as_of_date / nav_cum）")
    upload = st.file_uploader("选择净值 Excel", type=["xlsx"], key="nav_uploader")
    default_as_of = st.date_input("文件未含 as_of_date 时使用", value=date.today(), key="nav_default_date")
    if upload is not None and st.button("写入 SQLite", key="nav_save"):
        inbox = resolve_inbox_dir()
        inbox.mkdir(parents=True, exist_ok=True)
        target = inbox / upload.name
        target.write_bytes(upload.getvalue())
        from ..ingest.client_nav import import_nav_xlsx

        try:
            rows = import_nav_xlsx(target, as_of_date=default_as_of.isoformat())
            st.success(f"导入完成，写入 {rows} 行；源文件保存在 {target}")
        except Exception as exc:
            st.error(f"导入失败：{exc}")
