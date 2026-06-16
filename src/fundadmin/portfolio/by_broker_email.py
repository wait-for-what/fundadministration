"""分券商组合持仓汇总邮件（全产品一封）。

用途:
- 从 build_product_reports() 的 results（每个含 holdings_by_broker 分券商明细）渲染
  一封 HTML 邮件：顶部全产品概览表 + 每产品一节分券商持仓表，跨券商持有的标的
  （同一只股票同时在中金/中信等多家券商）整行高亮。
- 供每日 sync-latest --with-email 在发送 matrix 汇总邮件后顺带发出。

约束:
- 纯 pandas 渲染，**不依赖 streamlit / 看板页**，可在 headless 定时任务中安全调用。
- 仅展示已配置 product_code 的产品（与结构层落地集合一致）。
"""

from __future__ import annotations

import logging
from datetime import date
from typing import Any

import pandas as pd

from fundadmin.clients.config import NAME_TO_PRODCODE, broker_label
from fundadmin.notifications.email import SmtpConfig, send_html_email
from fundadmin.portfolio.notifier import build_weight_matrix_table

# 顶部权重矩阵中不展示的产品（与 matrix 邮件口径一致）。
_MATRIX_EXCLUDE = {"沐泽1号"}

logger = logging.getLogger(__name__)

# 券商列展示顺序（其余券商排在后面）。
_BROKER_ORDER = ["中金", "中信", "申万宏源", "未知"]
# 每产品明细表最多展示的标的数（按合计市值降序）。
_CAP_PER_PRODUCT = 25


def _fmt_int(value: Any) -> str:
    try:
        return f"{int(round(float(value))):,}"
    except (TypeError, ValueError):
        return "0"


def _fmt_mv(value: Any) -> str:
    return f"{float(value):,.2f}" if pd.notna(value) else "-"


def _fmt_pct(value: Any) -> str:
    return f"{float(value)*100:.2f}%" if pd.notna(value) else "-"


def _per_stock_view(by_broker: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """把分券商明细（每行 = 标的×券商）透视成每标的一行、各券商股数成列。

    返回 (view, broker_cols)。view 列：名称/代码/<各券商股数>/合计股数/合计市值/权重/来源券商数。
    """
    df = by_broker.copy()
    df["broker_cn"] = df["broker"].map(broker_label)
    df["shares"] = pd.to_numeric(df["shares"], errors="coerce").fillna(0.0)
    df["market_value_cny"] = pd.to_numeric(df["market_value_cny"], errors="coerce")
    df["weight"] = pd.to_numeric(df["weight"], errors="coerce")

    meta = df.groupby("group_key", as_index=False).agg(
        名称=("company", "first"),
        代码=("ticker", "first"),
        合计市值=("market_value_cny", "sum"),
        权重=("weight", "sum"),
    )
    shares = df.pivot_table(
        index="group_key", columns="broker_cn", values="shares", aggfunc="sum"
    ).fillna(0.0)

    present = list(shares.columns)
    ordered = [b for b in _BROKER_ORDER if b in present]
    rest = sorted([b for b in present if b not in ordered], key=lambda c: -float(shares[c].sum()))
    broker_cols = ordered + rest
    shares = shares[broker_cols]

    n_brokers = (shares != 0).sum(axis=1)
    total_shares = shares.sum(axis=1)
    shares = shares.reset_index()
    shares["合计股数"] = total_shares.values
    shares["来源券商数"] = n_brokers.values

    view = meta.merge(shares, on="group_key").drop(columns=["group_key"])
    view = view.sort_values("合计市值", ascending=False, na_position="last").reset_index(drop=True)
    view = view[["名称", "代码", *broker_cols, "合计股数", "合计市值", "权重", "来源券商数"]]
    return view, broker_cols


# --- 布局样式常量 ---
_FONT = "-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif"
_C_BORDER = "#e5e7eb"
_C_HEADBAR = "#26425f"   # 概览表头（深蓝）
_C_BAND = "#eef4fb"      # 卡片头/合计行（浅蓝）
_C_CROSS = "#fff7e0"     # 跨券商行（浅黄）
_C_CROSS_ACCENT = "#e0a800"
_C_ZEBRA = "#f7f9fb"     # 斑马纹


def _render_product_card(
    *, product_name: str, product_code: str, as_of: str, by_broker: pd.DataFrame
) -> tuple[str, dict[str, Any]]:
    """单产品卡片：头部色带（产品名 + 关键指标）+ 分券商持仓表（斑马纹 + 跨券商高亮）。"""
    view, broker_cols = _per_stock_view(by_broker)
    total_mv = float(pd.to_numeric(by_broker["market_value_cny"], errors="coerce").sum())
    n_stocks = len(view)
    n_cross = int((view["来源券商数"] > 1).sum())

    by_b = (
        by_broker.assign(
            b=by_broker["broker"].map(broker_label),
            mv=pd.to_numeric(by_broker["market_value_cny"], errors="coerce"),
        )
        .groupby("b", as_index=False)["mv"].sum()
        .sort_values("mv", ascending=False)
    )
    chips = "　".join(f"{r.b} ¥{r.mv:,.0f}" for r in by_b.itertuples() if pd.notna(r.mv))

    # 头部色带（用 table 实现左右分栏，邮件客户端兼容性优于 flex）。
    header = (
        f"<table style='width:100%;border-collapse:collapse;background:{_C_BAND}'><tr>"
        f"<td style='padding:11px 14px;font-size:15px;font-weight:700;color:#1f2937'>{product_name}"
        f"<span style='font-weight:500;color:#6b7280;font-size:12px'>　{product_code} · {as_of}</span></td>"
        f"<td style='padding:11px 14px;text-align:right;font-size:12px;color:#4b5563;white-space:nowrap'>"
        f"¥{total_mv:,.0f}　·　{n_stocks} 只　·　跨券商 <b style='color:#b8860b'>{n_cross}</b></td>"
        f"</tr></table>"
    )
    chip_bar = (
        f"<div style='padding:7px 14px;font-size:12px;color:#6b7280;"
        f"border-bottom:1px solid #f0f0f0'>各券商市值：{chips}</div>"
    )

    cols = [("名称", "left"), ("代码", "left"), *[(b, "right") for b in broker_cols],
            ("合计股数", "right"), ("合计市值", "right"), ("权重", "right")]
    th = "".join(
        f"<th style='padding:7px 10px;text-align:{a};font-size:12px;font-weight:600;"
        f"color:#374151;background:#f1f5f9;border-bottom:2px solid #dbe3ec'>{c}</th>"
        for c, a in cols
    )

    shown = view.head(_CAP_PER_PRODUCT)
    truncated = len(view) - len(shown)
    rows: list[str] = []
    for i, (_, r) in enumerate(shown.iterrows()):
        cross = r["来源券商数"] > 1
        if cross:
            bg = _C_CROSS
            accent = f"box-shadow:inset 3px 0 0 {_C_CROSS_ACCENT};"
        else:
            bg = "#ffffff" if i % 2 == 0 else _C_ZEBRA
            accent = ""
        cell = f"padding:6px 10px;border-bottom:1px solid #eef1f4;background:{bg};"
        tds = [
            f"<td style='{cell}{accent}'>{r['名称']}</td>",
            f"<td style='{cell}font-weight:600'>{r['代码']}{' 🔗' if cross else ''}</td>",
        ]
        for b in broker_cols:
            v = int(round(float(r[b]))) if pd.notna(r[b]) else 0
            stl = (
                "color:#1a73e8;font-weight:600" if (cross and v)
                else ("color:#c7ccd1" if v == 0 else "color:#374151")
            )
            tds.append(f"<td style='{cell}text-align:right;{stl}'>{_fmt_int(v)}</td>")
        tds.append(f"<td style='{cell}text-align:right;font-weight:600'>{_fmt_int(r['合计股数'])}</td>")
        tds.append(f"<td style='{cell}text-align:right'>{_fmt_mv(r['合计市值'])}</td>")
        tds.append(f"<td style='{cell}text-align:right'>{_fmt_pct(r['权重'])}</td>")
        rows.append("<tr>" + "".join(tds) + "</tr>")

    note = (
        f"<div style='padding:6px 12px;color:#9ca3af;font-size:12px'>"
        f"共 {len(view)} 只，显示前 {_CAP_PER_PRODUCT}（按市值）</div>"
        if truncated > 0 else ""
    )
    table = (
        f"<table style='border-collapse:collapse;width:100%;font-size:12.5px'>"
        f"<thead><tr>{th}</tr></thead><tbody>{''.join(rows)}</tbody></table>"
    )
    card = (
        f"<div style='border:1px solid {_C_BORDER};border-radius:10px;overflow:hidden;"
        f"margin:0 0 16px;box-shadow:0 1px 2px rgba(0,0,0,0.04)'>"
        f"{header}{chip_bar}{table}{note}</div>"
    )
    summary = {
        "product_name": product_name,
        "product_code": product_code,
        "as_of": as_of,
        "n_stocks": n_stocks,
        "n_cross": n_cross,
        "total_mv": total_mv,
    }
    return card, summary


def build_by_broker_summary_html(
    results: list[dict[str, Any]], *, trade_date: date
) -> str | None:
    """渲染全产品分券商持仓汇总 HTML；无任何可展示产品时返回 None。"""
    cards: list[str] = []
    overview: list[dict[str, Any]] = []
    for r in results:
        pname = str(r.get("product_name", "") or "")
        pcode = NAME_TO_PRODCODE.get(pname)
        if not pcode:
            # 与结构层一致：仅展示已配置 product_code 的产品。
            continue
        by_broker = r.get("holdings_by_broker")
        if by_broker is None or getattr(by_broker, "empty", True):
            continue
        as_of = str(r.get("trade_date", "") or trade_date.isoformat())
        card, summary = _render_product_card(
            product_name=pname, product_code=pcode, as_of=as_of, by_broker=by_broker
        )
        cards.append(card)
        overview.append(summary)

    if not overview:
        return None

    total_mv_all = sum(s["total_mv"] for s in overview)
    total_cross_all = sum(s["n_cross"] for s in overview)

    kpi = (
        f'<div style="margin:2px 0 14px;color:#374151;font-size:13px">'
        f'{len(overview)} 只产品　·　合计市值 <b>¥{total_mv_all:,.0f}</b>　·　'
        f'跨券商标的合计 <b style="color:#b8860b">{total_cross_all}</b> 行</div>'
    )
    legend = (
        f'<div style="margin:0 0 16px;padding:8px 12px;background:{_C_CROSS};'
        f'border:1px solid #f3e2b3;border-radius:8px;font-size:12px;color:#7a5b00">'
        f'<span style="display:inline-block;width:10px;height:10px;background:{_C_CROSS};'
        f'border:1px solid {_C_CROSS_ACCENT};border-radius:2px;vertical-align:middle"></span>'
        ' 黄色行 / 🔗 = 同一标的同时在多家券商持有；各券商股数已分列，合计股数与合计市值为并表口径。</div>'
    )
    section_label = (
        'margin:18px 0 8px;font-size:13px;font-weight:600;color:#374151;'
        f'border-left:3px solid {_C_HEADBAR};padding-left:8px'
    )

    matrix_table = build_weight_matrix_table(
        results, trade_date, exclude_products=_MATRIX_EXCLUDE
    )

    return (
        f'<div style="font-family:{_FONT};color:#1f2937;max-width:880px;margin:0 auto">'
        '<h2 style="margin:0 0 2px;font-size:20px">组合持仓 · 分券商来源</h2>'
        f'<div style="color:#6b7280;font-size:13px">全产品汇总 · {trade_date.isoformat()}</div>'
        f'{kpi}'
        f'<div style="{section_label}">持仓权重矩阵</div>'
        f'{matrix_table}'
        f'<div style="{section_label}">分券商明细</div>'
        f'{legend}'
        f'{"".join(cards)}'
        '<p style="color:#9ca3af;font-size:12px;margin-top:14px">'
        '数据源自每日券商报表自动构建（中金 / 中信 等）。</p>'
        "</div>"
    )


def send_by_broker_summary_email(
    results: list[dict[str, Any]],
    *,
    trade_date: date,
    smtp_config: SmtpConfig,
    to_addrs: list[str],
) -> bool:
    """构建并发送分券商持仓汇总邮件。

    返回 True 表示已发送；无可展示数据返回 False（不发空邮件）。
    异常向上抛出，由调用方决定是否容错。
    """
    html = build_by_broker_summary_html(results, trade_date=trade_date)
    if html is None:
        return False
    total_cross = sum(
        int((_per_stock_view(r["holdings_by_broker"])[0]["来源券商数"] > 1).sum())
        for r in results
        if NAME_TO_PRODCODE.get(str(r.get("product_name", "") or ""))
        and r.get("holdings_by_broker") is not None
        and not getattr(r.get("holdings_by_broker"), "empty", True)
    )
    subject = f"【组合持仓·分券商汇总】{trade_date.isoformat()}（跨券商 {total_cross} 行）"
    send_html_email(smtp_config, subject=subject, html_body=html, to_addrs=to_addrs)
    return True


__all__ = ["build_by_broker_summary_html", "send_by_broker_summary_email"]
