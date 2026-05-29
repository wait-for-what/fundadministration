"""日志初始化。

用途:
- 提供 get_logger() 统一获取带基础格式的 logger，供 CLI / jobs 复用。

边界:
- 仅做最小配置；不接管第三方库 root logger 的高级配置。
"""

from __future__ import annotations

import logging
import os

_CONFIGURED = False


def _configure_root() -> None:
    global _CONFIGURED
    if _CONFIGURED:
        return
    level = os.getenv("LOG_LEVEL", "INFO").upper()
    logging.basicConfig(
        level=getattr(logging, level, logging.INFO),
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    _CONFIGURED = True


def get_logger(name: str) -> logging.Logger:
    """返回按模块命名的 logger（首次调用时做一次 basicConfig）。"""
    _configure_root()
    return logging.getLogger(name)
