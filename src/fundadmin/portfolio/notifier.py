"""基金持仓报告邮件通知。

用途:
- 组装 HTML 汇总邮件并发送。

输入:
- results: build_product_reports() 的返回列表。
- trade_date: 交易日。
- chart_paths: dict[product_name, Path] 图表路径映射。
- smtp_config: SmtpConfig。
- to_addrs: 收件人列表。

输出:
- 发送 HTML 邮件（含内联 PNG 图表）。

失败行为:
- SMTP 配置缺失或收件人为空时抛出 ValueError。
- 邮件发送失败时抛出 EmailSendError 及其子类。

调用示例:
- `send_portfolio_email(results, trade_date=date(2026,4,17), chart_paths={...}, smtp_config=cfg, to_addrs=["a@b.com"])`
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Any

from fundadmin.notifications.email import SmtpConfig, send_html_email

# 默认不在邮件中展示持仓明细的产品
_DEFAULT_EXCLUDE_FROM_CHARTS: set[str] = {"沐泽1号"}


def build_email_html(
    results: list[dict[str, Any]],
    trade_date: date,
    chart_paths: dict[str, Path],
    *,
    exclude_products: set[str] | None = None,
) -> str:
    """组装 HTML 邮件正文。

    包含：
    1. 汇总表格（产品名、单位净值、资产净值、持仓数量、总市值）。
    2. 每只产品的饼图（通过 CID 引用内联图片），可排除指定产品。
    3. 报告文件路径附注。

    参数:
        exclude_products: 不展示持仓分布图的产品名称集合。
    """
    exclude = exclude_products or set()

    rows: list[str] = []
    for r in results:
        pname = r.get("product_name", "")
        unit_nav = r.get("unit_nav")
        asset_nav = r.get("asset_nav")
        nav = r.get("nav")
        total = r.get("total_holdings", 0)
        total_mv = r.get("total_market_value_cny")

        unit_nav_str = f"{unit_nav:,.4f}" if unit_nav is not None else "N/A"
        asset_nav_str = f"{asset_nav:,.2f}" if asset_nav is not None else "N/A"
        nav_str = f"{nav:,.2f}" if nav is not None else "N/A"
        mv_str = f"{total_mv:,.2f}" if total_mv is not None else "N/A"

        rows.append(
            f"<tr>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;'>{pname}</td>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{unit_nav_str}</td>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{asset_nav_str}</td>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{nav_str}</td>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{total}</td>"
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{mv_str}</td>"
            f"</tr>"
        )

    # 图表区域：排除指定产品
    chart_sections: list[str] = []
    for r in results:
        pname = r.get("product_name", "")
        if pname in exclude:
            continue
        chart_path = chart_paths.get(pname)
        if chart_path and chart_path.exists():
            cid = f"chart_{_safe_cid(pname)}"
            chart_sections.append(
                f"<div style='margin:24px 0;text-align:center;'>"
                f"<h3 style='font-size:14px;color:#333;margin-bottom:8px;'>{pname}</h3>"
                f"<img src='cid:{cid}' style='max-width:600px;width:100%;height:auto;border:1px solid #eee;border-radius:4px;' />"
                f"</div>"
            )

    # 报告文件路径附注
    file_notes: list[str] = []
    for r in results:
        pname = r.get("product_name", "")
        out_xlsx = r.get("out_xlsx", "")
        if out_xlsx:
            file_notes.append(f"<li>{pname}: {out_xlsx}</li>")

    exclude_note = ""
    if exclude:
        exclude_note = f"<p style='color:#888;font-size:12px;'>注：{', '.join(sorted(exclude))} 的持仓明细未在邮件中展示。</p>"

    html = f"""<!DOCTYPE html>
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
<h2>📊 基金持仓汇总 — {trade_date.isoformat()}</h2>
<table>
<thead>
<tr>
<th style="text-align:left;">产品</th>
<th style="text-align:right;">单位净值</th>
<th style="text-align:right;">资产净值</th>
<th style="text-align:right;">净值(权重)</th>
<th style="text-align:right;">持仓数量</th>
<th style="text-align:right;">总市值(CNY)</th>
</tr>
</thead>
<tbody>
{''.join(rows)}
</tbody>
</table>
{exclude_note}

<h2>📈 持仓分布</h2>
{''.join(chart_sections)}

<div class="footer">
<p><strong>报告文件：</strong></p>
<ul style="margin:4px 0;padding-left:18px;">
{''.join(file_notes)}
</ul>
<p style="margin-top:12px;">本邮件由 MarketAnalysis 自动生成。</p>
</div>
</body>
</html>"""
    return html


def _build_single_product_html(
    r: dict[str, Any],
    trade_date: date,
    chart_paths: dict[str, Path],
) -> str:
    """为单只产品组装 HTML 邮件正文。"""
    pname = r.get("product_name", "")
    unit_nav = r.get("unit_nav")
    asset_nav = r.get("asset_nav")
    nav = r.get("nav")
    total = r.get("total_holdings", 0)
    total_mv = r.get("total_market_value_cny")

    unit_nav_str = f"{unit_nav:,.4f}" if unit_nav is not None else "N/A"
    asset_nav_str = f"{asset_nav:,.2f}" if asset_nav is not None else "N/A"
    nav_str = f"{nav:,.2f}" if nav is not None else "N/A"
    mv_str = f"{total_mv:,.2f}" if total_mv is not None else "N/A"

    chart_section = ""
    chart_path = chart_paths.get(pname)
    if chart_path and chart_path.exists():
        cid = f"chart_{_safe_cid(pname)}"
        chart_section = (
            f"<div style='margin:24px 0;text-align:center;'>"
            f"<img src='cid:{cid}' style='max-width:600px;width:100%;height:auto;border:1px solid #eee;border-radius:4px;' />"
            f"</div>"
        )

    out_xlsx = r.get("out_xlsx", "")
    file_note = f"<li>{pname}: {out_xlsx}</li>" if out_xlsx else ""

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
<h2>📊 {pname} — {trade_date.isoformat()}</h2>
<table>
<thead>
<tr>
<th style="text-align:left;">指标</th>
<th style="text-align:right;">数值</th>
</tr>
</thead>
<tbody>
<tr><td style='padding:8px 12px;border:1px solid #ddd;'>单位净值</td><td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{unit_nav_str}</td></tr>
<tr><td style='padding:8px 12px;border:1px solid #ddd;'>资产净值</td><td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{asset_nav_str}</td></tr>
<tr><td style='padding:8px 12px;border:1px solid #ddd;'>净值(权重)</td><td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{nav_str}</td></tr>
<tr><td style='padding:8px 12px;border:1px solid #ddd;'>持仓数量</td><td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{total}</td></tr>
<tr><td style='padding:8px 12px;border:1px solid #ddd;'>总市值(CNY)</td><td style='padding:8px 12px;border:1px solid #ddd;text-align:right;'>{mv_str}</td></tr>
</tbody>
</table>

<h2>📈 持仓分布</h2>
{chart_section}

<div class="footer">
<p><strong>报告文件：</strong></p>
<ul style="margin:4px 0;padding-left:18px;">
{file_note}
</ul>
<p style="margin-top:12px;">本邮件由 MarketAnalysis 自动生成。</p>
</div>
</body>
</html>"""


def send_per_product_emails(
    results: list[dict[str, Any]],
    *,
    trade_date: date,
    chart_paths: dict[str, Path],
    smtp_config: SmtpConfig,
    to_addrs: list[str],
    exclude_products: set[str] | None = None,
) -> None:
    """为每只产品（除排除列表外）分别发送邮件。

    每封邮件包含：
    - 该产品的汇总信息表格
    - 持仓分布饼图（内联）
    - 产品报告 Excel 附件

    参数:
        results: 产品报告结果列表。
        trade_date: 交易日。
        chart_paths: 产品名到图表路径的映射。
        smtp_config: SMTP 配置。
        to_addrs: 收件人地址列表。
        exclude_products: 不发送持仓明细的产品名称集合，默认排除 "沐泽1号"。
    """
    if not to_addrs:
        raise ValueError("to_addrs is empty")

    exclude = exclude_products if exclude_products is not None else _DEFAULT_EXCLUDE_FROM_CHARTS

    for r in results:
        pname = r.get("product_name", "")
        if pname in exclude:
            continue

        html_body = _build_single_product_html(r, trade_date, chart_paths)

        # 内联图片
        images: dict[str, bytes] = {}
        chart_path = chart_paths.get(pname)
        if chart_path and chart_path.exists():
            cid = f"chart_{_safe_cid(pname)}"
            images[cid] = chart_path.read_bytes()

        # 附件：产品报告 Excel
        attachments: dict[str, bytes] = {}
        out_xlsx = r.get("out_xlsx", "")
        if out_xlsx:
            xlsx_path = Path(out_xlsx)
            if xlsx_path.exists():
                attachments[xlsx_path.name] = xlsx_path.read_bytes()

        subject = f"[基金持仓] {trade_date.isoformat()} {pname} 持仓报告"

        try:
            send_html_email(
                smtp_config,
                subject=subject,
                html_body=html_body,
                to_addrs=to_addrs,
                images=images if images else None,
                attachments=attachments if attachments else None,
            )
        except Exception as exc:
            # 单只产品发送失败不影响其他产品
            import warnings
            warnings.warn(f"{pname} 邮件发送失败: {exc}", stacklevel=2)


def _safe_cid(name: str) -> str:
    """将产品名转换为安全的 CID 标识符。"""
    import re
    return re.sub(r"[^a-zA-Z0-9\u4e00-\u9fff]", "_", name)


# ---- 持仓权重矩阵邮件 ----


def build_weight_matrix_html(
    results: list[dict[str, Any]],
    trade_date: date,
    chart_paths: dict[str, Path],
    *,
    exclude_products: set[str] | None = None,
) -> str:
    """生成持仓权重矩阵 HTML 邮件正文。

    表格结构：
    - 第一行表头：标的名称 | 产品A | 产品B | ...
    - 第一列：所有标的（去重，按字母序）
    - 单元格：权重百分比（如 12.34%），无持仓留空
    """
    import pandas as pd

    exclude = exclude_products or set()

    # 收集每个产品的 ticker -> weight 映射
    product_weights: dict[str, dict[str, float]] = {}
    all_tickers: set[str] = set()

    for r in results:
        pname = r.get("product_name", "")
        if pname in exclude:
            continue
        holdings = r.get("holdings_raw")
        if holdings is None or holdings.empty:
            product_weights[pname] = {}
            continue
        df = holdings.copy()
        # 使用 company 作为显示名称（ticker 可能为空）
        df["display_name"] = df["group_key"].fillna(df["company"]).fillna("未知")
        weights = {}
        for _, row in df.iterrows():
            name = str(row["display_name"]).strip()
            weight = row.get("weight")
            if name and pd.notna(weight):
                weights[name] = float(weight)
                all_tickers.add(name)
        product_weights[pname] = weights

    product_names = sorted(product_weights.keys())

    if not all_tickers or not product_names:
        return f"""<!DOCTYPE html>
<html><body><p>{trade_date.isoformat()} 无有效持仓数据。</p></body></html>"""

    # 按第一个产品的权重降序排列标的；权重为 None 的排最后（按字母序）
    first_product = product_names[0]
    ticker_list = sorted(
        all_tickers,
        key=lambda t: (product_weights[first_product].get(t) or -1),
        reverse=True,
    )

    # 构建 HTML 表格行
    header_cols = ["<th style='background:#4a90d9;color:#fff;padding:10px 12px;border:1px solid #ddd;text-align:center;'>标的名称</th>"]
    for pname in product_names:
        header_cols.append(f"<th style='background:#4a90d9;color:#fff;padding:10px 12px;border:1px solid #ddd;text-align:center;white-space:nowrap;'>{pname}</th>")
    header_row = "<tr>" + "".join(header_cols) + "</tr>"

    body_rows: list[str] = []
    for i, ticker in enumerate(ticker_list):
        bg = "#f9fafc" if i % 2 == 0 else "#ffffff"
        cells = [f"<td style='padding:8px 12px;border:1px solid #ddd;background:{bg};font-weight:600;'>{ticker}</td>"]
        for pname in product_names:
            w = product_weights[pname].get(ticker)
            if w is not None:
                cell = f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;background:{bg};'>{w:.2%}</td>"
            else:
                cell = f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:center;background:{bg};color:#ccc;'>—</td>"
            cells.append(cell)
        body_rows.append("<tr>" + "".join(cells) + "</tr>")

    # 总权重行
    total_cells = ["<td style='padding:8px 12px;border:1px solid #ddd;background:#e8f0fe;font-weight:700;'>总权重</td>"]
    for pname in product_names:
        total_w = sum(product_weights[pname].values())
        total_cells.append(
            f"<td style='padding:8px 12px;border:1px solid #ddd;text-align:right;background:#e8f0fe;font-weight:700;'>{total_w:.2%}</td>"
        )
    body_rows.append("<tr>" + "".join(total_cells) + "</tr>")

    # 图表区域（排除产品）
    chart_sections: list[str] = []
    for pname in product_names:
        chart_path = chart_paths.get(pname)
        if chart_path and chart_path.exists():
            cid = f"chart_{_safe_cid(pname)}"
            chart_sections.append(
                f"<div style='margin:20px 0;text-align:center;'>"
                f"<h4 style='font-size:13px;color:#333;margin-bottom:6px;'>{pname}</h4>"
                f"<img src='cid:{cid}' style='max-width:500px;width:100%;height:auto;border:1px solid #eee;border-radius:4px;' />"
                f"</div>"
            )

    # 产品净值汇总行
    nav_rows: list[str] = []
    for label, key, fmt in [("单位净值", "unit_nav", ":,.4f"), ("资产净值", "asset_nav", ":,.2f")]:
        cells = [f"<td style='padding:6px 10px;border:1px solid #ddd;background:#f0f4f8;font-weight:600;'>{label}</td>"]
        for pname in product_names:
            r = next((x for x in results if x.get("product_name") == pname), {})
            val = r.get(key)
            val_str = format(val, fmt[1:]) if val is not None else "N/A"
            cells.append(f"<td style='padding:6px 10px;border:1px solid #ddd;text-align:right;background:#f0f4f8;'>{val_str}</td>")
        nav_rows.append("<tr>" + "".join(cells) + "</tr>")

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{ font-family: "Segoe UI", "Microsoft YaHei", sans-serif; color: #333; line-height: 1.6; max-width: 1200px; margin: 0 auto; padding: 20px; }}
h2 {{ font-size: 18px; color: #1a1a1a; border-bottom: 2px solid #4a90d9; padding-bottom: 8px; margin-top: 28px; }}
.matrix-table {{ border-collapse: collapse; width: 100%; margin: 12px 0; font-size: 12px; }}
.matrix-table th {{ padding: 10px 12px; border: 1px solid #ddd; text-align: center; }}
.matrix-table td {{ padding: 6px 10px; border: 1px solid #ddd; }}
.footer {{ margin-top: 32px; padding-top: 16px; border-top: 1px solid #eee; font-size: 12px; color: #888; }}
</style>
</head>
<body>
<h2>📊 持仓权重矩阵 — {trade_date.isoformat()}</h2>
<table class="matrix-table">
<thead>
{header_row}
{''.join(nav_rows)}
</thead>
<tbody>
{''.join(body_rows)}
</tbody>
</table>

<h2>📈 持仓分布</h2>
{''.join(chart_sections)}

<div class="footer">
<p style="margin-top:12px;">本邮件由 MarketAnalysis 自动生成。</p>
</div>
</body>
</html>"""
    return html


def send_matrix_email(
    results: list[dict[str, Any]],
    *,
    trade_date: date,
    chart_paths: dict[str, Path],
    smtp_config: SmtpConfig,
    to_addrs: list[str],
    attachments: dict[str, bytes] | None = None,
    exclude_products: set[str] | None = None,
) -> None:
    """发送持仓权重矩阵汇总邮件。

    邮件包含：
    - 产品 × 标的 的持仓权重矩阵表格
    - 各产品的饼图（内联，排除指定产品）
    - 可选附件
    """
    if not to_addrs:
        raise ValueError("to_addrs is empty")

    exclude = exclude_products if exclude_products is not None else _DEFAULT_EXCLUDE_FROM_CHARTS
    html_body = build_weight_matrix_html(results, trade_date, chart_paths, exclude_products=exclude)

    # 读取图表文件为内联图片
    images: dict[str, bytes] = {}
    for r in results:
        pname = r.get("product_name", "")
        if pname in exclude:
            continue
        chart_path = chart_paths.get(pname)
        if chart_path and chart_path.exists():
            cid = f"chart_{_safe_cid(pname)}"
            images[cid] = chart_path.read_bytes()

    product_count = len([r for r in results if r.get("product_name", "") not in exclude])
    subject = f"[基金持仓] {trade_date.isoformat()} 跨券商持仓权重矩阵（{product_count}只产品）"

    send_html_email(
        smtp_config,
        subject=subject,
        html_body=html_body,
        to_addrs=to_addrs,
        images=images if images else None,
        attachments=attachments if attachments else None,
    )
