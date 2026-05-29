"""客户/净值/持仓的配置常量与路径解析。

用途:
- 解析 FUND_DB_URL / FUND_INBOX_DIR / FUND_OUTPUTS_DIR 等环境变量。
- 集中维护产品代码与名称映射，供 ingest / compute / pages 复用。

输入:
- 环境变量 FUND_DB_URL（必需），其他为可选。

输出:
- 字符串路径与映射字典。

失败行为:
- FUND_DB_URL 未设置时，调用 resolve_db_url 抛 RuntimeError。
"""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path

from fundadmin.core.config import get_env

# 产品代码 -> 名称（与 portfolio.client_notifier._PRODCODE_TO_NAME 对齐）
PRODCODE_TO_NAME: Mapping[str, str] = {
    "SCD704": "全球视野",
    "SCY282": "种子",
    "SLL384": "铂金1号",
    "SNJ280": "铂金2号",
    "SXQ602": "铂金8号",
}


def resolve_db_url() -> str:
    """读取 FUND_DB_URL；缺失时报错。

    约束:
    - 客户/净值/持仓走自包含数据库（默认 SQLite），不回退到其他 DB_URL。
    """
    url = get_env("FUND_DB_URL", required=True)
    return str(url).strip()


def resolve_inbox_dir() -> Path:
    """Excel/CSV 落盘与上传目录。默认 outputs/inbox。"""
    raw = get_env("FUND_INBOX_DIR")
    if raw:
        return Path(raw).expanduser().resolve()
    from fundadmin.core.paths import repo_root

    return (repo_root() / "outputs" / "inbox").resolve()


def resolve_outputs_dir() -> Path:
    """看板生成产物目录。默认 outputs/exports。"""
    raw = get_env("FUND_OUTPUTS_DIR")
    if raw:
        return Path(raw).expanduser().resolve()
    from fundadmin.core.paths import repo_root

    return (repo_root() / "outputs" / "exports").resolve()
