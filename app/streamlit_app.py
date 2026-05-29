"""FundAdministration 内网看板（Streamlit）。

用途:
- 两个标签页：产品持仓 / 客户净值。数据来自本地 SQLite（FUND_DB_URL）。

调用入口:
- streamlit run app/streamlit_app.py

边界:
- 只读展示；写入由 `fundadmin clients import-*` 与 `fundadmin portfolio ...` 负责。
"""

from __future__ import annotations

import streamlit as st

from fundadmin.clients.pages import client_nav_page, portfolio_page
from fundadmin.clients.schema import init_db

st.set_page_config(page_title="FundAdministration", layout="wide")
init_db()

tab_holdings, tab_nav = st.tabs(["产品持仓", "客户净值"])
with tab_holdings:
    portfolio_page.render()
with tab_nav:
    client_nav_page.render()
