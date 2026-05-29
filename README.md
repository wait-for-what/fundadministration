# FundAdministration

基金运营系统：客户/投资人台账、产品净值、券商持仓、客户分产品净值通知。

本仓库从 **MarketAnalysis** 剥离而来，目标是与市场研究代码彻底隔离，独立部署、独立交付：
- 处理真实投资人 PII 与客户资金，访问控制 / 数据库 / 备份与 MarketAnalysis 完全独立
- 不依赖 `marketanalysis` 包；仅复制了极小的基础设施切片（config / db engine / email）
- 数据底座默认是**自包含 SQLite**（便于备份与交接：拷一个文件即可）

> Provenance: 代码迁移自 MarketAnalysis（`domains/fund_portfolio` + `domains/internal_dashboard` 的基金部分）。源提交记录见本仓库初始提交说明。watchlist（重仓价格区间）功能在迁移时已废弃，不包含在内。

## 安装

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e .          # 如需从 MySQL 导入客户清单：pip install -e ".[mysql]"
cp .env.example .env      # 填写 FUND_DB_URL / SMTP_* / IMAP_*
```

## 常用命令

```bash
python -m fundadmin.cli --help
python -m fundadmin.cli clients init-db
python -m fundadmin.cli clients import-clients --csv path/to/clients.csv
python -m fundadmin.cli clients import-nav --xlsx path/to/summary.xlsx
python -m fundadmin.cli clients import-portfolio --xlsx path/to/holdings.xlsx
python -m fundadmin.cli portfolio email-sync
python -m fundadmin.cli portfolio notify-client-nav --dry-run

# 内网看板（产品持仓 / 客户净值）
streamlit run app/streamlit_app.py
```

## 结构

```
src/fundadmin/
  core/           # config(get_env) / paths / logging
  db/engine.py    # get_engine（FUND_DB_URL）
  notifications/  # email.py（SMTP）
  clients/        # 客户/净值/持仓：schema(3表) / store / ingest / compute / pages
  portfolio/      # 券商持仓 Excel 抓取解析 + 客户净值通知 + 按公司汇总
app/              # Streamlit 看板
deploy/launchd/   # Mac mini 定时任务 + cloudflared 示例
```
