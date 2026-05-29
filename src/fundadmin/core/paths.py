"""仓库路径解析。

用途:
- 提供 repo_root() 作为产物目录（outputs / data / inbox）的锚点。

实现:
- 从本文件向上查找含 pyproject.toml 的目录；找不到时回退到固定层级。
"""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path


@lru_cache(maxsize=1)
def repo_root() -> Path:
    here = Path(__file__).resolve()
    for parent in (here, *here.parents):
        if (parent / "pyproject.toml").exists():
            return parent
    # 回退：src/fundadmin/core/paths.py -> 仓库根
    return here.parents[3]
