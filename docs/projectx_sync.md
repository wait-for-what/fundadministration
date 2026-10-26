# projectx 发布：基金持仓与产品净值

把本地库中的**基金持仓**与**产品净值**完整发布到 `projectx` MySQL 库（本机
`127.0.0.1:3306/projectx`），供其他系统（如 MarketAnalysis 侧的量化分析）直接读取。

## 数据流

```
券商邮件(IMAP) --backfill_history.py--> 本地 SQLite (FUND_DB_URL)
                                          fund_positions / product_valuation
                                                |
                                      fundadmin projectx sync（按估值日整快照镜像）
                                                v
                                     projectx MySQL
                                   fund_positions / fund_product_valuation
```

- 本地库是**唯一事实源**：projectx 里的表是派生副本，可随时删掉重建。
- projectx 连接串：环境变量 `PROJECTX_DB_URL`（`mysql+pymysql://.../projectx?charset=utf8mb4`）；
  未设置时回退 `DB_URL` 并把 schema 改写为 `projectx`（与 `scripts/export_x_cust_rpt.py` 约定一致）。

## 表方案

命名加 `fund_` 前缀做命名空间隔离（projectx 里已有 `x_cust_rpt`、`etf_rotation_universe`）。
列与本地表一一对应；`utf8mb4`；数值统一 `DOUBLE`；时间戳列 `updated_at` 记录同步时间
（`ON UPDATE CURRENT_TIMESTAMP`，本地 `loaded_at` 不外发）。

### fund_positions —— 分券商持仓明细

对应本地 `fund_positions`（核心结构层）。每行 = 某估值日某产品在某券商的某标的（含保证金/成本/损益超集列）。

主键：`(as_of_date, product_code, broker, ticker, instrument_name, contract_no)`
索引：`(product_code, as_of_date)`

| 列 | 类型 | 说明 |
| --- | --- | --- |
| as_of_date | DATE, PK | 估值日 |
| product_code / product_name | VARCHAR(32)/(128), /PK | 产品代码/名称 |
| broker | VARCHAR(32), PK, DEFAULT '' | 券商（cicc/citic/swhysc；空串=未分券商） |
| ticker / instrument_name | VARCHAR(64)/(255), PK, DEFAULT '' | 标的代码/名称（现金行允许空串） |
| asset_class / direction / currency | VARCHAR(32)/(16)/(8) | 资产类别 / 方向 / 币种 |
| fx_rate / quantity | DOUBLE | 汇率 / 数量 |
| market_value_cny / market_value_local | DOUBLE | 市值（人民币 / 原币） |
| weight | DOUBLE | 占净值比 |
| cost_price_local / cost_value_local / cost_ccy | DOUBLE/DOUBLE/VARCHAR(8) | 成本价 / 成本值 / 成本币种 |
| financing_cost / init_margin / maint_margin / mtm_pnl | DOUBLE | 融资成本 / 初始保证金 / 维持保证金 / 当日盈亏 |
| contract_no | VARCHAR(64), PK, DEFAULT '' | 合约号 |
| source_files | TEXT | 来源附件（逗号分隔文件名） |
| ingest_id | BIGINT | 本地附件入库 ID（raw 层回溯用） |
| updated_at | DATETIME | 同步时间（自动） |

### fund_product_valuation —— 产品级净值/规模

对应本地 `product_valuation`。每行 = 某估值日某产品的净值与规模汇总，是"产品净值"的主出口。

主键：`(as_of_date, product_code)`；索引：`(product_code, as_of_date)`

| 列 | 类型 | 说明 |
| --- | --- | --- |
| as_of_date | DATE, PK | 估值日 |
| product_code / product_name | VARCHAR(32)/(128) | 产品代码/名称 |
| unit_nav | DOUBLE | 单位净值 |
| asset_nav | DOUBLE | 资产净值 |
| nav_for_weight | DOUBLE | 计权重用净值 |
| total_holdings | INT | 持仓标的数 |
| total_market_value_cny | DOUBLE | 持仓总市值（人民币） |
| ingest_id | BIGINT | 本地附件入库 ID |
| updated_at | DATETIME | 同步时间（自动） |

### 设计决策

- **按估值日整快照替换（DELETE+INSERT，单事务）**而非 upsert：本地每日构建就是整快照
  （`replace_positions` 同款语义），镜像端用同样语义天然幂等，本地删除/修正能完整传播，
  且不会因主键列演化留下脏行。
- **不发布** `fund_portfolio_holdings`（旧表，已被 fund_positions 取代且为空）、
  `fund_transactions`（当前为空，待成交流水积累后再加 `fund_transactions` 表）、
  `attachment_ingest`/`raw_sheet_rows`（原始无损层体积大且仅供本地排障）。
- 单位净值以 `fund_product_valuation.unit_nav` 为准；本地 `product_nav_history`
  （summary Excel 导入的累计净值流）一旦有数据，可按同模式加 `fund_product_nav` 表。

## 使用

```bash
# 1) 建表（幂等；sync 内也会自动建）
python -m fundadmin.cli projectx init-db

# 2) 全量镜像（默认全部估值日；远端本地已不存在的日期会被清理）
python -m fundadmin.cli projectx sync
python -m fundadmin.cli projectx sync --dates 2026-08-18,2026-08-19   # 只动指定日期

# 3) 对账（按表按估值日比行数；不一致退出码 1）
python -m fundadmin.cli projectx verify
```

## 完整历史回填（过往数据）

历史数据的源头是券商邮件。完整回填两步（均幂等、不发任何客户邮件）：

```bash
# 1) 从邮箱按收件日下载附件（message-id 去重，可反复执行）
python scripts/backfill_history.py --start 2026-01-01 --download-only

# 2) 用已下载 inbox 按交易日重建入库（删旧重写）
python scripts/backfill_history.py --build-only
```

然后执行 `projectx sync` 即可把全部历史镜像进 projectx。建议把 `sync` 挂在每日
`sync-latest` 之后，让 projectx 每天自动跟进最新估值日。
