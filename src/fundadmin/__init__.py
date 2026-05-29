"""FundAdministration 基金运营系统。

用途:
- 客户/投资人台账、产品净值、券商持仓与客户分产品净值通知。

边界:
- 自包含系统，不依赖 marketanalysis 包；数据底座默认独立 SQLite（FUND_DB_URL）。
"""

__version__ = "0.1.0"
