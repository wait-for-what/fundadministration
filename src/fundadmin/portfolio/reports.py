"""报表/邮件完整性检查 — 不和 IMAP/CLI/DB 耦合，只处理已构建的数据。

治理提示::
    从 operations.py 拆出（god module 治理 2026-04 P2.7）。
    不引入 operations.py 的任何函数，保持单向依赖。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from fundadmin.portfolio.cross_broker_report import PRODUCT_CONFIG


def product_email_completion_issues(
    *,
    results: list[dict[str, Any]],
    summary_path: Path | None,
    require_charts: bool,
    chart_paths: dict[str, Path],
) -> list[str]:
    issues: list[str] = []
    expected_names = [str(cfg.get("name", "")).strip() for cfg in PRODUCT_CONFIG]
    expected_names = [name for name in expected_names if name]
    by_name = {
        str(result.get("product_name", "")).strip(): result
        for result in results
        if str(result.get("product_name", "")).strip()
    }

    missing_products = [name for name in expected_names if name not in by_name]
    if missing_products:
        issues.append(f"missing product reports: {', '.join(missing_products)}")

    missing_nav = [name for name in expected_names if name in by_name and by_name[name].get("nav") is None]
    if missing_nav:
        issues.append(f"missing NAV: {', '.join(missing_nav)}")

    empty_holdings = [
        name
        for name in expected_names
        if name in by_name and int(by_name[name].get("total_holdings") or 0) <= 0
    ]
    if empty_holdings:
        issues.append(f"empty holdings: {', '.join(empty_holdings)}")

    underpriced: list[str] = []
    for name in expected_names:
        result = by_name.get(name)
        if result is None:
            continue
        hr = result.get("holdings_raw")
        if hr is None or getattr(hr, "empty", True) or "market_value_cny" not in getattr(hr, "columns", []):
            continue
        if pd.to_numeric(hr["market_value_cny"], errors="coerce").isna().any():
            underpriced.append(name)
    if underpriced:
        issues.append(f"holdings missing market value: {', '.join(underpriced)}")

    missing_outputs: list[str] = []
    for name in expected_names:
        result = by_name.get(name)
        if result is None:
            continue
        out_xlsx = str(result.get("out_xlsx", "") or "").strip()
        if not out_xlsx or not Path(out_xlsx).exists():
            missing_outputs.append(name)
    if missing_outputs:
        issues.append(f"missing product Excel outputs: {', '.join(missing_outputs)}")

    if summary_path is None or not summary_path.exists():
        issues.append("missing summary Excel output")

    if require_charts:
        missing_charts = [
            name
            for name in expected_names
            if name in by_name and (
                chart_paths.get(name) is None or not chart_paths[name].exists()
            )
        ]
        if missing_charts:
            issues.append(f"missing charts: {', '.join(missing_charts)}")

    return issues


def client_unit_nav_missing(results: list[dict[str, Any]]) -> list[str]:
    expected = [str(cfg.get("name", "")).strip() for cfg in PRODUCT_CONFIG]
    expected = [name for name in expected if name]
    by_name = {
        str(r.get("product_name", "")).strip(): r
        for r in results
        if str(r.get("product_name", "")).strip()
    }
    return [name for name in expected if name in by_name and by_name[name].get("unit_nav") is None]
