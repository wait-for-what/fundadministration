"""客户分产品净值通知系统。

用途:
- 从 projectx.x_cust_rpt 读取客户持仓信息。
- 结合 fund portfolio 净值数据，为每个客户生成个性化净值通知邮件。

输入:
- trade_date: 交易日。
- nav_source: 净值数据来源（summary Excel 路径或 build_product_reports 结果）。
- smtp_config: SMTP 配置。

输出:
- 向有 email 地址的客户发送净值通知邮件。

调用示例:
    send_client_nav_emails(
        trade_date=date(2026, 4, 17),
        summary_xlsx=Path(".../fund_portfolio_summary_2026-04-17.xlsx"),
        smtp_config=SmtpConfig(...),
    )
"""

from __future__ import annotations

import warnings
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

from fundadmin.notifications.email import SmtpConfig, send_html_email

# prodcode -> 产品名称（与 fund portfolio 系统一致）
_PRODCODE_TO_NAME: dict[str, str] = {
    "SCD704": "全球视野",
    "SCY282": "种子",
    "SLL384": "铂金1号",
    "SNJ280": "铂金2号",
    "SXQ602": "铂金8号",
}


def load_clients(trade_date: date) -> pd.DataFrame:
    """从本地 clients 表加载客户持仓数据。

    用途:
    - 客户清单是本仓库 SQLite（FUND_DB_URL）的单一事实来源，由
      `fundadmin clients import-clients` 维护。历史上来自 MySQL projectx.x_cust_rpt，
      剥离后改读本地表以保持自包含。

    返回:
    - DataFrame，列：custname, prodcode, prodname, 持有份额, email, mobile, product_name。
      仅保留有 email 的有效客户；trade_date 仅用于调用方语义，不参与过滤。
    """
    from fundadmin.clients.store import load_clients as _load_clients_table

    df = _load_clients_table(active_only=True)
    if df.empty:
        return df
    df = df[df["email"].notna() & (df["email"].astype(str).str.strip() != "")].copy()
    if df.empty:
        return df
    df = df.rename(columns={"holding_shares": "持有份额"})
    df["product_name"] = df["prodcode"].map(_PRODCODE_TO_NAME)
    df["持有份额"] = pd.to_numeric(df["持有份额"], errors="coerce")
    return df


def load_nav_map(summary_xlsx: Path) -> dict[str, dict[str, Any]]:
    """从 summary Excel 加载产品净值映射。

    返回 {product_name: {"unit_nav": float, "asset_nav": float}}
    """
    if not summary_xlsx.exists():
        raise FileNotFoundError(f"Summary Excel not found: {summary_xlsx}")

    df = pd.read_excel(summary_xlsx, engine="openpyxl")
    nav_map: dict[str, dict[str, Any]] = {}
    for _, row in df.iterrows():
        pname = str(row.get("product_name", "")).strip()
        unit_nav = row.get("unit_nav")
        asset_nav = row.get("asset_nav")
        if pname:
            nav_map[pname] = {
                "unit_nav": float(unit_nav) if pd.notna(unit_nav) else None,
                "asset_nav": float(asset_nav) if pd.notna(asset_nav) else None,
            }
    return nav_map


def _build_client_html(
    custname: str,
    holdings: list[dict[str, Any]],
    trade_date: date,
) -> str:
    """为单个客户组装 HTML 邮件正文。"""
    rows: list[str] = []
    total_mv: float = 0.0

    for h in holdings:
        pname = h["product_name"]
        unit_nav = h.get("unit_nav")
        shares = h.get("shares")

        unit_nav_str = f"{unit_nav:,.4f}" if unit_nav is not None else "N/A"
        shares_str = f"{shares:,.2f}" if shares is not None else "N/A"

        mv = None
        if unit_nav is not None and shares is not None:
            mv = unit_nav * shares
            total_mv += mv
            mv_str = f"{mv:,.2f}"
        else:
            mv_str = "N/A"

        rows.append(
            f"<tr>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;'>{pname}</td>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{unit_nav_str}</td>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{shares_str}</td>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{mv_str}</td>"
            f"</tr>"
        )

    total_mv_str = f"{total_mv:,.2f}" if total_mv > 0 else "N/A"

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{ font-family: "Segoe UI", "Microsoft YaHei", sans-serif; color: #333; line-height: 1.6; max-width: 800px; margin: 0 auto; padding: 20px; }}
h2 {{ font-size: 18px; color: #1a1a1a; border-bottom: 2px solid #4a90d9; padding-bottom: 8px; margin-top: 28px; }}
table {{ border-collapse: collapse; width: 100%; margin: 12px 0; font-size: 13px; }}
th {{ background: #f5f7fa; color: #333; font-weight: 600; text-align: left; padding: 10px 12px; border: 1px solid #ddd; }}
td {{ padding: 8px 12px; border: 1px solid #ddd; }}
.footer {{ margin-top: 32px; padding-top: 16px; border-top: 1px solid #eee; font-size: 12px; color: #888; }}
</style>
</head>
<body>
<h2>📊 弘运盛泰产品净值通知 — {trade_date.isoformat()}</h2>
<p>尊敬的 <strong>{custname}</strong>：</p>
<p>您持有的产品净值信息如下：</p>
<table>
<thead>
<tr>
<th style="text-align:left;">产品名称</th>
<th style="text-align:right;">单位净值</th>
<th style="text-align:right;">持有份额</th>
<th style="text-align:right;">持仓市值（估算）</th>
</tr>
</thead>
<tbody>
{''.join(rows)}
<tr style="background:#e8f0fe;font-weight:700;">
<td style="padding:8px 12px;border:1px solid #ddd;">合计</td>
<td style="padding:8px 12px;border:1px solid #ddd;"></td>
<td style="padding:8px 12px;border:1px solid #ddd;"></td>
<td style="padding:8px 12px;border:1px solid #ddd;text-align:right;">{total_mv_str}</td>
</tr>
</tbody>
</table>
<div class="footer">
<p>注：持仓市值 = 持有份额 × 单位净值，仅供参考，实际以托管估值为准。</p>
<p style="margin-top:8px;">如您对净值、持有份额等信息有任何疑问，可直接回复本邮件或致电联系我们。</p>
<p style="margin-top:12px;">本邮件由弘运盛泰自动生成。</p>
</div>
</body>
</html>"""


def send_client_nav_emails(
    trade_date: date,
    summary_xlsx: Path,
    smtp_config: SmtpConfig,
) -> dict[str, Any]:
    """发送客户净值通知邮件。

    参数:
        trade_date: 交易日。
        summary_xlsx: fund_portfolio_summary 的 Excel 路径。
        smtp_config: SMTP 配置。

    返回:
        发送统计信息字典。
    """
    clients = load_clients(trade_date)
    if clients.empty:
        return {"sent": 0, "skipped": 0, "reason": "no clients with email"}

    nav_map = load_nav_map(summary_xlsx)

    # 按客户分组
    sent = 0
    skipped = 0
    for custname, group in clients.groupby("custname"):
        email = group["email"].iloc[0]
        if not email or "@" not in email:
            skipped += 1
            continue

        holdings: list[dict[str, Any]] = []
        for _, row in group.iterrows():
            pname = row.get("product_name")
            if not pname:
                continue
            nav = nav_map.get(pname, {})
            holdings.append({
                "product_name": pname,
                "unit_nav": nav.get("unit_nav"),
                "asset_nav": nav.get("asset_nav"),
                "shares": row.get("持有份额"),
            })

        if not holdings:
            skipped += 1
            continue

        html_body = _build_client_html(custname, holdings, trade_date)
        subject = f"[净值通知] {trade_date.isoformat()} 弘运盛泰产品净值更新"

        try:
            send_html_email(
                smtp_config,
                subject=subject,
                html_body=html_body,
                to_addrs=[email],
            )
            sent += 1
            print(f"[OK] Sent to {custname} ({email})")
        except Exception as exc:
            warnings.warn(f"Email send failed for {custname}: {exc}", stacklevel=2)
            skipped += 1

    return {"sent": sent, "skipped": skipped, "total": len(clients.groupby("custname").size())}
