"""客户导入。

用途:
- import_clients_csv: 从 projectx.x_cust_rpt 一次性导出的 CSV 落 clients 表。
  期望列：custname, prodcode, prodname, 持有份额(或 holding_shares), email, mobile。

输入:
- 文件路径。

输出:
- 写入行数。

失败行为:
- 列缺失时抛 ValueError；CSV 编码异常由 read_csv_robust 报错。
"""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd
from sqlalchemy.engine import Engine

from fundadmin.portfolio.parsers.common import read_csv_robust

from ..config import PRODCODE_TO_NAME
from ..store import upsert_clients

logger = logging.getLogger(__name__)


def import_clients_csv(csv_path: str | Path, *, engine: Engine | None = None) -> int:
    """从导出的 CSV 写入 clients。"""
    path = Path(csv_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    df = read_csv_robust(path)
    if df.empty:
        return 0
    df = df.rename(columns={c: str(c).strip() for c in df.columns})

    rename_map: dict[str, str] = {}
    if "持有份额" in df.columns and "holding_shares" not in df.columns:
        rename_map["持有份额"] = "holding_shares"
    df = df.rename(columns=rename_map)

    for col in ("custname", "prodcode"):
        if col not in df.columns:
            raise ValueError(f"clients CSV 缺少必要列: {col}")
    for col in ("prodname", "holding_shares", "email", "mobile", "active"):
        if col not in df.columns:
            df[col] = None

    df["custname"] = df["custname"].astype(str).str.strip()
    df["prodcode"] = df["prodcode"].astype(str).str.strip()
    df = df[df["custname"].astype(bool) & df["prodcode"].astype(bool)]
    if df.empty:
        return 0

    df["holding_shares"] = pd.to_numeric(df["holding_shares"], errors="coerce")
    n_bad_shares = int(df["holding_shares"].isna().sum())
    if n_bad_shares:
        logger.warning(
            "import_clients_csv(%s): %d 行 holding_shares 非数值/缺失，将以空值入库（clients-store-6）",
            path.name, n_bad_shares,
        )
    df["active"] = pd.to_numeric(df["active"], errors="coerce").fillna(1).astype(int)
    df["prodname"] = df["prodname"].where(
        df["prodname"].astype(str).str.strip().ne("") & df["prodname"].notna(),
        df["prodcode"].map(PRODCODE_TO_NAME),
    )

    return upsert_clients(df, engine=engine)


__all__ = ["import_clients_csv"]
