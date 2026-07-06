#!/usr/bin/env bash
# 用途: 定时增量同步券商持仓邮件并构建产品报表（launchd 调用入口）。
# 透传给 `fundadmin portfolio sync-latest`，额外参数原样传入。
set -euo pipefail
cd "$(dirname "$0")/.."

# 收紧本地密钥文件权限，避免同机其它账户读到 IMAP/SMTP 口令 (security-1)。
if [ -f .env ]; then
  chmod 600 .env 2>/dev/null || true
fi

# 单实例锁：避免 launchd 定时任务与手动运行重叠，损坏 JSON 去重状态
# 或重复处理邮件 (reliability-1)。macOS 不自带 flock，用原子 mkdir 实现。
LOCK_DIR="${TMPDIR:-/tmp}/fundadmin-portfolio-sync.lock"
if ! mkdir "$LOCK_DIR" 2>/dev/null; then
  lock_pid="$(cat "$LOCK_DIR/pid" 2>/dev/null || true)"
  if [ -n "$lock_pid" ] && kill -0 "$lock_pid" 2>/dev/null; then
    echo "[run_portfolio_sync] 另一个同步进程 (pid $lock_pid) 正在运行，跳过本次。" >&2
    exit 0
  fi
  echo "[run_portfolio_sync] 清理陈旧锁 $LOCK_DIR（持有进程已不存在）。" >&2
  rm -rf "$LOCK_DIR"
  mkdir "$LOCK_DIR"
fi
echo "$$" > "$LOCK_DIR/pid"
# 注意: 不能用 exec，否则 trap 不会在 python 退出后触发清理锁目录。
trap 'rm -rf "$LOCK_DIR"' EXIT

if [ -d .venv ]; then
  # shellcheck disable=SC1091
  source .venv/bin/activate
fi
python -m fundadmin.cli portfolio sync-latest "$@"
