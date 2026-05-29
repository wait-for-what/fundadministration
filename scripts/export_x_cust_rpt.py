"""一次性脚本：从 MySQL projectx.x_cust_rpt 导出客户清单为 CSV。

用途:
- 在能连 MySQL 的机器上执行，把 x_cust_rpt 导成 CSV，供
  `fundadmin clients import-clients --csv` 导入本地 SQLite。
- 输出 CSV 含 PII，不入库。

输入:
- 环境变量 PROJECTX_DB_URL（mysql+pymysql://.../projectx），或 DB_URL（自动把 schema 改写到 projectx）。
- --out CSV 路径，默认 outputs/inbox/clients.csv。

输出:
- CSV 列: custname, prodcode, prodname, holding_shares, email, mobile

调用示例:
- python scripts/export_x_cust_rpt.py --out /tmp/clients.csv
- python -m fundadmin.cli clients import-clients --csv /tmp/clients.csv

依赖:
- 需要 pymysql：pip install -e ".[mysql]"
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine


def _resolve_projectx_url() -> str:
    url = os.getenv("PROJECTX_DB_URL")
    if url:
        return url.strip()
    raw = os.getenv("DB_URL")
    if not raw:
        raise SystemExit("缺少连接串：请设置 PROJECTX_DB_URL 或 DB_URL")
    raw = raw.replace("charset=utf8mb4f", "charset=utf8mb4")
    base = raw.rsplit("/", 1)[0]
    return f"{base}/projectx?charset=utf8mb4"


def main() -> int:
    ap = argparse.ArgumentParser(description="Export projectx.x_cust_rpt to CSV")
    ap.add_argument("--out", type=Path, default=Path("outputs/inbox/clients.csv"))
    args = ap.parse_args()

    sql = (
        "SELECT custname, prodcode, prodname, `持有份额` AS holding_shares, email, mobile "
        "FROM x_cust_rpt"
    )
    engine = create_engine(_resolve_projectx_url())
    df = pd.read_sql(sql, engine)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(args.out, index=False, encoding="utf-8-sig")
    print(f"exported {len(df)} rows -> {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
