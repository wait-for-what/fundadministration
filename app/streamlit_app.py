"""FundAdministration 内网看板（Streamlit）。

用途:
- 产品持仓展示。数据来自本地 SQLite（FUND_DB_URL）。
- 客户净值页已随 product_nav_history 下线移除；产品净值以 product_valuation 为准。

调用入口:
- streamlit run app/streamlit_app.py

边界:
- 只读展示；写入由 `fundadmin clients import-*` 与 `fundadmin portfolio ...` 负责。
"""

from __future__ import annotations

import streamlit as st

from fundadmin.clients.pages import portfolio_page
from fundadmin.clients.schema import init_db

st.set_page_config(page_title="FundAdministration", layout="wide")
init_db()

portfolio_page.render()
