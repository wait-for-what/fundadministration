#!/usr/bin/env bash
# 用途: 启动 FundAdministration 内网看板（Streamlit）。launchd / 手工均可用。
set -euo pipefail
cd "$(dirname "$0")/.."
if [ -d .venv ]; then
  # shellcheck disable=SC1091
  source .venv/bin/activate
fi
exec streamlit run app/streamlit_app.py \
  --server.port "${FUND_DASHBOARD_PORT:-8520}" \
  --server.headless true
