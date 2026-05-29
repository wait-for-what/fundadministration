#!/usr/bin/env bash
# 用途: macOS / Linux 手工执行 fundadmin CLI 的薄封装。
set -euo pipefail
cd "$(dirname "$0")/.."
if [ -d .venv ]; then
  # shellcheck disable=SC1091
  source .venv/bin/activate
fi
exec python -m fundadmin.cli "$@"
