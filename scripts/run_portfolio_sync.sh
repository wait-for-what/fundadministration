#!/usr/bin/env bash
# 用途: 定时增量同步券商持仓邮件并构建产品报表（launchd 调用入口）。
# 透传给 `fundadmin portfolio sync-latest`，额外参数原样传入。
set -euo pipefail
cd "$(dirname "$0")/.."
if [ -d .venv ]; then
  # shellcheck disable=SC1091
  source .venv/bin/activate
fi
exec python -m fundadmin.cli portfolio sync-latest "$@"
