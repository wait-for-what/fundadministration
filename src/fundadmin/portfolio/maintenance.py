"""fund_inbox 滚动清理工具。

用途:
- 在 `outputs/excels/fund_inbox/` 下按交易日组织的子目录里，删除超过保留窗口的
  旧目录，控制本地磁盘占用。
- 仅作用于 `fund_inbox/`，**不动 `outputs/reports/fund_portfolios/`**
  （报表为最终交付物，由用户手动管理）。

输入:
- `inbox_root`: 默认 `outputs/excels/fund_inbox/`。
- `keep_last`: 保留的最近交易日子目录数量（按目录名 ISO 日期降序排序后保留前 N 个）。
- `dry_run`: True 时只列出待删除项，不实际删除。

输出:
- `PruneResult`: 包含保留与删除的目录列表、释放空间统计。

失败行为:
- `inbox_root` 不存在时直接返回空结果，不抛错。
- 删除某个子目录失败时记录 warning 并跳过，不中断整次清理。

调用示例:
- `python -m apps.cli ops fund-portfolio prune-inbox --keep-last 30 --dry-run`
- `prune_inbox(inbox_root=Path("outputs/excels/fund_inbox"), keep_last=30)`
"""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path

_DATE_DIR_PATTERN = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")


@dataclass(frozen=True)
class PruneResult:
    """清理操作结果。"""

    inbox_root: Path
    keep_last: int
    dry_run: bool
    kept_dirs: tuple[Path, ...] = field(default_factory=tuple)
    pruned_dirs: tuple[Path, ...] = field(default_factory=tuple)
    skipped_dirs: tuple[Path, ...] = field(default_factory=tuple)
    bytes_pruned: int = 0
    failures: tuple[tuple[Path, str], ...] = field(default_factory=tuple)


def _is_date_dir(path: Path) -> bool:
    if not path.is_dir():
        return False
    return bool(_DATE_DIR_PATTERN.match(path.name))


def _dir_size_bytes(path: Path) -> int:
    """递归求和目录下所有常规文件大小；权限或符号链接异常时忽略。"""
    total = 0
    for entry in path.rglob("*"):
        try:
            if entry.is_file():
                total += entry.stat().st_size
        except OSError:
            # 权限/链接异常：跳过，不影响整体统计
            continue
    return total


def prune_inbox(
    *,
    inbox_root: Path,
    keep_last: int,
    dry_run: bool = False,
) -> PruneResult:
    """清理 fund_inbox 下超过保留窗口的日期子目录。

    边界:
    - `keep_last <= 0` 视为非法，抛 ValueError。调用方应当显式给保留数。
    - 非日期命名（`_state/`、临时目录、文件等）一律 skip，不动。
    - 即使整次只 dry-run，也会返回完整 `kept / pruned / bytes_pruned` 统计供
      CLI 友好打印。
    - 实际删除使用 `shutil.rmtree`，删除前先扣 `bytes_pruned` 计数。
    """
    if keep_last <= 0:
        raise ValueError("keep_last must be > 0")
    root = Path(inbox_root)
    if not root.exists():
        return PruneResult(inbox_root=root, keep_last=keep_last, dry_run=dry_run)

    date_dirs: list[Path] = []
    skipped: list[Path] = []
    for child in sorted(root.iterdir(), key=lambda p: p.name):
        if _is_date_dir(child):
            date_dirs.append(child)
        else:
            # `_state/`、非日期目录、零散文件保持原样
            skipped.append(child)

    # 按日期名降序，保留前 keep_last 个，其余为待删除
    date_dirs_sorted = sorted(date_dirs, key=lambda p: p.name, reverse=True)
    kept = date_dirs_sorted[:keep_last]
    candidates = date_dirs_sorted[keep_last:]

    pruned: list[Path] = []
    failures: list[tuple[Path, str]] = []
    bytes_pruned = 0
    for path in candidates:
        size = _dir_size_bytes(path)
        if dry_run:
            pruned.append(path)
            bytes_pruned += size
            continue
        try:
            shutil.rmtree(path)
        except OSError as exc:
            failures.append((path, str(exc)))
            continue
        pruned.append(path)
        bytes_pruned += size

    return PruneResult(
        inbox_root=root,
        keep_last=keep_last,
        dry_run=dry_run,
        kept_dirs=tuple(kept),
        pruned_dirs=tuple(pruned),
        skipped_dirs=tuple(skipped),
        bytes_pruned=bytes_pruned,
        failures=tuple(failures),
    )


def format_bytes(num_bytes: int) -> str:
    """人类可读字节单位（用于 CLI 输出）。"""
    value = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(value) < 1024.0 or unit == "TB":
            return f"{value:.1f} {unit}"
        value /= 1024.0
    return f"{value:.1f} TB"
