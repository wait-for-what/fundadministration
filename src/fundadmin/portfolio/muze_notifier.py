"""沐泽1号每日持仓变动邮件。

数据源:
    中金四级科目估值表（SQJ420）。估值表非逐笔成交披露，"持仓动作"由相邻
    估值日的持股数量差反推，金额按估值日价格估算（区间净额，不含费用）。

收件人:
    仅发送到 ``muze`` 分组（``EMAIL_GROUP_MUZE``）。该分组不回退 ``EMAIL_TO``，
    避免误发给通用收件人；未配置时不发送并返回 no_recipients。
"""

from __future__ import annotations

import html
import json
import logging
from datetime import date
from pathlib import Path
from typing import Any, Mapping

import pandas as pd

from fundadmin.core.config import get_env
from fundadmin.notifications.email import SmtpConfig, send_html_email
from fundadmin.notifications.recipients import resolve_email_group
from fundadmin.portfolio.parsers.muze_valuation import (
    PRODUCT_CODE,
    PRODUCT_NAME,
    daily_trade_flow,
    holdings_changes,
    parse_muze_dir,
)

logger = logging.getLogger(__name__)

_SMTP_PORT_DEFAULT = 465
_ACTION_COLOR = {
    "新进": "#c0392b",
    "加仓": "#e67e22",
    "减仓": "#2c7fb8",
    "清仓": "#27ae60",
}


def _wan(v: Any, digits: int = 1, signed: bool = False) -> str:
    try:
        val = float(v)
    except (TypeError, ValueError):
        return ""
    if pd.isna(val):
        return ""
    sign = "+" if signed else ""
    return f"{val / 1e4:{sign},.{digits}f}万"


def _pct(v: Any) -> str:
    try:
        val = float(v)
    except (TypeError, ValueError):
        return ""
    return "" if pd.isna(val) else f"{val:.1f}%"


def _num(v: Any) -> str:
    try:
        return f"{float(v):,.0f}"
    except (TypeError, ValueError):
        return ""


def _price(v: Any) -> str:
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return ""


def smtp_config_from_env(
    *,
    host: str = "",
    port: int = 0,
    user: str = "",
    password: str = "",
    from_addr: str = "",
) -> SmtpConfig:
    """从环境变量构造 SmtpConfig，显式参数优先。"""
    h = str(host or get_env("SMTP_HOST", default="") or "")
    p = int(port or get_env("SMTP_PORT", default="") or _SMTP_PORT_DEFAULT)
    u = str(user or get_env("SMTP_USER", default="") or "")
    pw = str(password or get_env("SMTP_PASS", default="") or "")
    frm = str(from_addr or get_env("EMAIL_FROM", default="") or u)
    if not h or not u or not pw:
        raise RuntimeError("SMTP 配置不完整：需要 SMTP_HOST/SMTP_USER/SMTP_PASS")
    return SmtpConfig(host=h, port=p, user=u, password=pw, from_addr=frm)


def _state_path() -> Path:
    from fundadmin.clients.config import resolve_outputs_dir

    return resolve_outputs_dir() / "state" / "muze_notify_state.json"


def _load_state(path: Path) -> dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_state(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def build_muze_daily_change_html(
    *,
    day_changes: pd.DataFrame,
    nav_row: Mapping[str, Any],
    holdings_today: pd.DataFrame,
    trade_date: str,
) -> str:
    """渲染单日持仓变动邮件 HTML。"""
    d = day_changes.copy()
    d["abs_amount"] = d["chg_amount"].abs()
    d = d.sort_values("abs_amount", ascending=False)

    buy_amount = float(d.loc[d["chg_amount"] > 0, "chg_amount"].sum())
    sell_amount = float(-d.loc[d["chg_amount"] < 0, "chg_amount"].sum())
    n_buy = int((d["chg_amount"] > 0).sum())
    n_sell = int((d["chg_amount"] < 0).sum())

    rows = []
    for r in d.itertuples(index=False):
        color = _ACTION_COLOR.get(str(r.action), "#555")
        rows.append(
            "<tr>"
            f"<td style='color:{color};font-weight:600'>{html.escape(str(r.action))}</td>"
            f"<td>{html.escape(str(r.ticker))}</td>"
            f"<td>{html.escape(str(r.instrument_name))}</td>"
            f"<td>{html.escape(str(r.market))}</td>"
            f"<td style='text-align:right'>{_num(r.prev_qty)}</td>"
            f"<td style='text-align:right'>{_num(r.quantity)}</td>"
            f"<td style='text-align:right;color:{color}'>{_num(r.qty_chg)}</td>"
            f"<td style='text-align:right'>{_price(r.price)}</td>"
            f"<td style='text-align:right;color:{color}'>{_wan(r.chg_amount, signed=True)}</td>"
            "</tr>"
        )

    hold = holdings_today.copy()
    hold = hold.sort_values("market_value", ascending=False)
    hold_total = float(hold["market_value"].sum()) or 1.0
    hold_rows = []
    for r in hold.itertuples(index=False):
        pnl_color = "#c0392b" if float(r.unrealized_pnl) >= 0 else "#27ae60"
        hold_rows.append(
            "<tr>"
            f"<td>{html.escape(str(r.ticker))}</td>"
            f"<td>{html.escape(str(r.instrument_name))}</td>"
            f"<td>{html.escape(str(r.market))}</td>"
            f"<td style='text-align:right'>{_num(r.quantity)}</td>"
            f"<td style='text-align:right'>{_wan(r.market_value)}</td>"
            f"<td style='text-align:right;color:{pnl_color}'>{_wan(r.unrealized_pnl, signed=True)}</td>"
            f"<td style='text-align:right'>{r.market_value / hold_total:.1%}</td>"
            "</tr>"
        )

    nav_tbl = (
        "<table><tbody>"
        f"<tr><th style='text-align:left'>单位净值</th><td>{float(nav_row.get('unit_nav') or 0):.4f}</td>"
        f"<th style='text-align:left'>资产净值</th><td>{_wan(nav_row.get('fund_asset_nav'), 1)}</td></tr>"
        f"<tr><th style='text-align:left'>股票仓位</th><td>{_pct(nav_row.get('stock_ratio_pct'))}</td>"
        f"<th style='text-align:left'>现金类仓位</th><td>{_pct(nav_row.get('cash_like_ratio_pct'))}</td></tr>"
        "</tbody></table>"
    )

    flow_tbl = (
        "<table><tbody><tr>"
        f"<th>买入额</th><td style='color:#c0392b;text-align:right'>{_wan(buy_amount)}</td>"
        f"<th>卖出额</th><td style='color:#27ae60;text-align:right'>{_wan(sell_amount)}</td>"
        f"<th>净买入</th><td style='text-align:right'>{_wan(buy_amount - sell_amount, signed=True)}</td>"
        "</tr><tr>"
        f"<th>买入笔数</th><td style='text-align:right'>{n_buy}</td>"
        f"<th>卖出笔数</th><td style='text-align:right'>{n_sell}</td>"
        f"<th>涉及标的</th><td style='text-align:right'>{d['ticker'].nunique()}</td>"
        "</tr></tbody></table>"
    )

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
body {{ font-family: "Segoe UI", "Microsoft YaHei", sans-serif; color:#333; line-height:1.6;
       max-width:860px; margin:0 auto; padding:20px; }}
h2 {{ font-size:18px; color:#1a1a1a; border-bottom:2px solid #4a90d9; padding-bottom:8px; }}
h3 {{ font-size:15px; color:#444; margin-top:24px; }}
table {{ border-collapse:collapse; width:100%; margin:10px 0; font-size:13px; }}
th {{ background:#f5f7fa; color:#333; font-weight:600; text-align:left; padding:8px 10px; border:1px solid #ddd; }}
td {{ padding:7px 10px; border:1px solid #ddd; }}
.footer {{ margin-top:28px; padding-top:14px; border-top:1px solid #eee; font-size:12px; color:#888; }}
.muted {{ color:#888; font-size:12px; }}
</style></head><body>
<h2>📈 {html.escape(PRODUCT_NAME)} · 每日持仓变动 — {html.escape(trade_date)}</h2>

<h3>组合概况</h3>
{nav_tbl}

<h3>当日交易动作（估值表推断）</h3>
{flow_tbl}

<h3>变动明细</h3>
<table>
<thead><tr><th>动作</th><th>代码</th><th>名称</th><th>市场</th>
<th>上期持股</th><th>当前持股</th><th>变动股数</th><th>价格</th><th>变动金额</th></tr></thead>
<tbody>{''.join(rows)}</tbody>
</table>

<h3>当前持仓</h3>
<table>
<thead><tr><th>代码</th><th>名称</th><th>市场</th><th>持股</th>
<th>市值</th><th>浮动盈亏</th><th>占股票组合</th></tr></thead>
<tbody>{''.join(hold_rows)}</tbody>
</table>

<div class="footer">
<p>数据来源：中金四级科目估值表（{html.escape(PRODUCT_CODE)}）。估值表非逐日逐笔披露，
持仓动作由相邻估值日持股数量差推算，变动金额按估值日价格估算，为区间净额，不含交易费用。</p>
<p>本邮件由弘运盛泰自动生成，仅发送至 muze 分组。</p>
</div>
</body></html>"""


def send_muze_daily_change_email(
    *,
    inbox_dir: str | Path,
    group: str = "muze",
    trade_date: str | None = None,
    smtp_config: SmtpConfig | None = None,
    state_file: str | Path | None = None,
    force: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """解析估值表并向 muze 分组发送最新估值日的持仓变动邮件。

    返回统计 dict：sent / reason / trade_date / recipients / n_changes / html(仅 dry-run)。
    无变动、已发送过该日、无收件人均不发送。
    """
    holdings, nav = parse_muze_dir(inbox_dir)
    if nav.empty or holdings.empty:
        return {"sent": False, "reason": "no_value_sheets"}

    target = str(trade_date or nav["as_of_date"].max())
    nav_row_df = nav[nav["as_of_date"] == target]
    if nav_row_df.empty:
        return {"sent": False, "reason": f"no_nav_for_{target}"}
    nav_row = nav_row_df.iloc[0].to_dict()

    changes = holdings_changes(holdings)
    day_changes = changes[changes["as_of_date"] == target].copy()
    if day_changes.empty:
        return {"sent": False, "reason": "no_changes", "trade_date": target}

    state_path = Path(state_file).expanduser() if state_file else _state_path()
    state = _load_state(state_path)
    if not force and state.get("last_sent_date") == target:
        return {"sent": False, "reason": "already_sent", "trade_date": target}

    recipients = resolve_email_group(group)
    if not recipients:
        return {
            "sent": False,
            "reason": f"no_recipients_for_group_{group}",
            "trade_date": target,
        }

    holdings_today = holdings[holdings["as_of_date"] == target].copy()
    html_body = build_muze_daily_change_html(
        day_changes=day_changes,
        nav_row=nav_row,
        holdings_today=holdings_today,
        trade_date=target,
    )
    subject = f"【{PRODUCT_NAME}·持仓变动】{target}"
    result: dict[str, Any] = {
        "sent": False,
        "reason": "dry_run" if dry_run else "",
        "trade_date": target,
        "recipients": recipients,
        "n_changes": int(len(day_changes)),
        "subject": subject,
    }
    if dry_run:
        result["html"] = html_body
        return result

    cfg = smtp_config or smtp_config_from_env()
    send_html_email(cfg, subject=subject, html_body=html_body, to_addrs=recipients)

    _save_state(
        state_path,
        {
            "last_sent_date": target,
            "last_sent_at": date.today().isoformat(),
            "recipients": recipients,
        },
    )
    result["sent"] = True
    result.pop("reason", None)
    logger.info("muze daily change email sent: %s -> %s", target, recipients)
    return result


__all__ = [
    "build_muze_daily_change_html",
    "send_muze_daily_change_email",
    "smtp_config_from_env",
]
