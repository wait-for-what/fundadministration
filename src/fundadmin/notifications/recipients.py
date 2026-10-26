"""邮件收件人分组解析。

约定:
- 分组名 -> 环境变量 ``EMAIL_GROUP_<GROUP>``（大写），逗号分隔地址。
- 仅 ``default`` 分组在未显式配置时回退到 ``EMAIL_TO``；其他分组不回退，
  避免误发给通用收件人。
- 沐泽1号专用分组为 ``muze``（``EMAIL_GROUP_MUZE``）。

示例 .env:
    EMAIL_GROUP_MUZE=someone@example.com
    EMAIL_GROUP_DEFAULT=ops@example.com,pm@example.com
    EMAIL_TO=legacy@example.com        # 仅在未配置 EMAIL_GROUP_DEFAULT 时生效
"""

from __future__ import annotations

from fundadmin.core.config import get_env


def resolve_email_group(group: str = "default") -> list[str]:
    """解析指定分组的收件人列表；仅 default 分组回退 EMAIL_TO，仍为空返回 []。"""
    name = str(group or "default").strip().lower()
    key = f"EMAIL_GROUP_{name.upper()}"
    raw = get_env(key, default="") or ""
    if not str(raw).strip() and name == "default":
        raw = get_env("EMAIL_TO", default="") or ""
    return [x.strip() for x in str(raw).split(",") if x.strip()]


__all__ = ["resolve_email_group"]
