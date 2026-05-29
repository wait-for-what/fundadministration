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

from dotenv import find_dotenv, load_dotenv

_LOADED = False


def load_env(dotenv_path: str | None = None) -> None:
    """加载 .env（仅一次）；未显式给路径时按当前工作目录向上查找。"""
    global _LOADED
    if not _LOADED:
        if dotenv_path:
            load_dotenv(dotenv_path=dotenv_path)
        else:
            env_path = find_dotenv(usecwd=True)
            load_dotenv(dotenv_path=env_path or None)
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
