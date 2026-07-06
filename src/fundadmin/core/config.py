"""环境变量与 .env 读取。

用途:
- 进程级一次性加载 .env，并提供统一的 get_env 读取入口。

输入:
- 环境变量 / 项目根 .env 文件。

输出:
- 字符串配置值；required=True 且缺失时抛 RuntimeError。
"""

from __future__ import annotations

import os
import stat
import warnings

from dotenv import find_dotenv, load_dotenv

_LOADED = False


def _warn_if_world_readable(env_path: str | None) -> None:
    """若 .env 对同组/其它用户可读，提示收紧权限（security-1）。

    .env 内含 IMAP/SMTP 明文口令，同机其它账户可读即等于凭据泄露。
    """
    if not env_path or not os.path.isfile(env_path):
        return
    try:
        mode = os.stat(env_path).st_mode
    except OSError:
        return
    if mode & (stat.S_IRGRP | stat.S_IROTH | stat.S_IWGRP | stat.S_IWOTH):
        warnings.warn(
            f".env 权限过宽（{oct(stat.S_IMODE(mode))}），含明文口令。"
            f"建议 `chmod 600 {env_path}`。",
            stacklevel=2,
        )


def load_env(dotenv_path: str | None = None) -> None:
    """加载 .env（仅一次）；未显式给路径时按当前工作目录向上查找。"""
    global _LOADED
    if not _LOADED:
        if dotenv_path:
            load_dotenv(dotenv_path=dotenv_path)
            _warn_if_world_readable(dotenv_path)
        else:
            env_path = find_dotenv(usecwd=True)
            load_dotenv(dotenv_path=env_path or None)
            _warn_if_world_readable(env_path or None)
        _LOADED = True


def get_env(name: str, default: str | None = None, required: bool = False) -> str | None:
    """读取环境变量；required 缺失时报错。"""
    load_env()
    val = os.getenv(name, default)
    if isinstance(val, str):
        val = val.strip()
    if required and not val:
        raise RuntimeError(f"{name} is not set in environment/.env")
    return val
