"""基金持仓可视化图表生成。

用途:
- 为每只产品生成持仓分布饼图（前10大持仓 + Others）。

输入:
- holdings_raw: pd.DataFrame（含 group_key, company, market_value_cny, weight, shares）。
- product_name, trade_date, nav, total_holdings: 图表元信息。

输出:
- PNG 文件写入指定路径。

调用示例:
- `generate_portfolio_pie_chart(df, product_name="铂金1号", trade_date=date(2026,4,17), nav=1.2345, total_holdings=25, out_path=Path("..."))`
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import matplotlib
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import pandas as pd

# 非交互式后端，避免在 headless 环境弹窗
matplotlib.use("Agg")

# 尝试配置中文字体（Windows / macOS / Linux 常见字体）
_CJK_FONTS = [
    "Microsoft YaHei",
    "SimHei",
    "PingFang SC",
    "Heiti SC",
    "WenQuanYi Micro Hei",
    "Noto Sans CJK SC",
    "Source Han Sans SC",
    "STHeiti",
    "Arial Unicode MS",
]

# 查找系统中可用的第一个中文字体
_CJK_FONT_NAME: str | None = None
_available_fonts = {f.name for f in fm.fontManager.ttflist}
for name in _CJK_FONTS:
    if name in _available_fonts:
        _CJK_FONT_NAME = name
        break

if _CJK_FONT_NAME:
    plt.rcParams["font.sans-serif"] = [_CJK_FONT_NAME] + plt.rcParams.get("font.sans-serif", [])
    plt.rcParams["axes.unicode_minus"] = False


def generate_portfolio_pie_chart(
    holdings_raw: pd.DataFrame,
    *,
    product_name: str,
    trade_date: date,
    nav: float | None,
    total_holdings: int,
    out_path: Path,
) -> Path:
    """生成持仓分布饼图并保存为 PNG。

    参数:
        holdings_raw: 未格式化的持仓 DataFrame，需含 `company` 和 `market_value_cny` 列。
        product_name: 产品名称，用于图表标题。
        trade_date: 交易日，用于图表副标题。
        nav: 净值，用于图表副标题。
        total_holdings: 持仓数量，用于图表副标题。
        out_path: 输出 PNG 文件路径。

    返回:
        写入后的文件路径。
    """
    if holdings_raw.empty or "market_value_cny" not in holdings_raw.columns:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        _generate_empty_chart(product_name, trade_date, out_path)
        return out_path

    # 取有效市值行，按市值降序排列
    df = (
        holdings_raw[holdings_raw["market_value_cny"].notna()]
        .copy()
        .sort_values("market_value_cny", ascending=False)
        .reset_index(drop=True)
    )
    if df.empty:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        _generate_empty_chart(product_name, trade_date, out_path)
        return out_path

    # 取前10，其余合并为 "其他"
    top_n = 10
    if len(df) > top_n:
        top = df.head(top_n).copy()
        others_value = df.iloc[top_n:]["market_value_cny"].sum()
        others = pd.DataFrame(
            [{"company": "其他", "market_value_cny": others_value}]
        )
        plot_df = pd.concat([top, others], ignore_index=True)
    else:
        plot_df = df.copy()

    labels = plot_df["company"].astype(str).tolist()
    values = plot_df["market_value_cny"].astype(float).tolist()
    total = sum(values)

    # 百分比标签：只显示占比 >= 2% 的切片
    def _pct_label(pct: float) -> str:
        return f"{pct:.1f}%" if pct >= 2.0 else ""

    fig, ax = plt.subplots(figsize=(8, 6), dpi=120)

    colors = plt.cm.Set3(range(len(labels)))

    wedges, texts, autotexts = ax.pie(
        values,
        labels=labels,
        autopct=_pct_label,
        startangle=90,
        colors=colors,
        pctdistance=0.75,
        labeldistance=1.08,
        textprops={"fontsize": 9},
    )

    for autotext in autotexts:
        autotext.set_fontsize(8)
        autotext.set_color("#333333")

    # 标题与副标题
    title = f"{product_name} 持仓分布"
    nav_str = f"NAV={nav:,.4f}" if nav is not None else "NAV=N/A"
    subtitle = f"{trade_date.isoformat()} | {nav_str} | 共{total_holdings}只标的"

    ax.set_title(title, fontsize=14, fontweight="bold", pad=10)
    fig.text(0.5, 0.92, subtitle, ha="center", fontsize=10, color="#666666")

    # 在图表右下方添加图例（公司名 + 市值）
    legend_labels = [
        f"{lbl}  {val:,.0f} ({val/total*100:.1f}%)" if total > 0 else lbl
        for lbl, val in zip(labels, values, strict=False)
    ]
    ax.legend(
        wedges,
        legend_labels,
        title="持仓明细",
        loc="center left",
        bbox_to_anchor=(1.02, 0, 0.5, 1),
        fontsize=8,
        title_fontsize=9,
    )

    plt.tight_layout()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    return out_path


def _generate_empty_chart(product_name: str, trade_date: date, out_path: Path) -> None:
    """生成无数据的占位图。"""
    fig, ax = plt.subplots(figsize=(6, 4), dpi=120)
    ax.text(
        0.5,
        0.5,
        f"{product_name}\n{trade_date.isoformat()}\n无持仓数据",
        ha="center",
        va="center",
        fontsize=14,
        color="#999999",
    )
    ax.axis("off")
    fig.savefig(out_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def generate_adjusted_nav_chart(
    nav_df: pd.DataFrame,
    *,
    product_code: str,
    product_name: str,
    events: pd.DataFrame | None = None,
    out_path: Path,
) -> Path:
    """绘制单只产品的净值走势对比：未复权 / 算术加回 / 分红再投资。

    输入:
        nav_df: 含 val_date, unit_nav, adj_nav（算术加回）, reinvest_nav（分红再投资）
            的 DataFrame，可由 adjusted_nav.build_nav_variants 得到。
        product_code / product_name: 用于标题与文件名。
        events: detect_dividend_events 的输出；在除息日画竖线并标注每份分红。
        out_path: PNG 输出路径。

    输出:
        PNG 文件路径。
    """
    if nav_df is None or nav_df.empty:
        _generate_empty_chart(product_name, date.today(), out_path)
        return out_path

    work = nav_df.sort_values("val_date")
    x = work["val_date"]
    raw = work["unit_nav"].astype(float)
    adj = work["adj_nav"].astype(float)
    rei = work["reinvest_nav"].astype(float) if "reinvest_nav" in work else None

    fig, ax = plt.subplots(figsize=(13, 5.5), dpi=150)

    if rei is not None:
        ax.plot(x, rei, color="#1f77b4", linewidth=1.9, label="分红再投资（推荐口径）")
        ax.plot(x, adj, color="#2ca02c", linewidth=1.1, linestyle="-.",
                label="算术加回（累计净值）")
    else:
        ax.plot(x, adj, color="#1f77b4", linewidth=1.9, label="复权净值（后复权）")
    ax.plot(x, raw, color="#ff7f0e", linewidth=1.1, linestyle="--", label="单位净值（未复权）")

    n_events = 0
    if events is not None and not events.empty:
        hits = events[events["product_code"] == product_code].sort_values("val_date")
        nav_by_date = work.set_index("val_date")["unit_nav"]
        for row in hits.itertuples():
            ax.axvline(row.val_date, color="#d62728", linewidth=0.8, alpha=0.45, linestyle=":")
            level = nav_by_date.get(row.val_date)
            if level is not None and pd.notna(level):
                ax.annotate(
                    f"-{row.amount_per_unit:g}",
                    xy=(row.val_date, float(level)),
                    fontsize=7,
                    color="#d62728",
                    rotation=90,
                    ha="right",
                    va="top",
                )
            n_events += 1

    last = work.iloc[-1]
    base = work.iloc[0]
    def _ret(series):
        s = series.astype(float)
        return float(s.iloc[-1] / s.iloc[0] - 1.0)

    raw_ret = _ret(raw)
    adj_ret = _ret(adj)
    rei_ret = _ret(rei) if rei is not None else None
    cum_div = float(work["cum_payout"].max()) if "cum_payout" in work else 0.0

    ax.set_title(f"{product_name}（{product_code}）净值走势 · 复权口径对比", fontsize=14, fontweight="bold")
    parts = [
        f"{pd.Timestamp(base['val_date']).date()} ~ {pd.Timestamp(last['val_date']).date()}",
        f"未复权 {raw_ret:+.2%}",
        f"算术加回 {adj_ret:+.2%}",
    ]
    if rei_ret is not None:
        parts.insert(2, f"分红再投资 {rei_ret:+.2%}")
    parts.append(f"累计每份分红 {cum_div:.4f}")
    parts.append(f"分红 {n_events} 次")
    fig.text(0.5, 0.90, " | ".join(parts), ha="center", fontsize=9, color="#666666")

    ax.set_ylabel("净值（元）")
    ax.grid(alpha=0.25, linewidth=0.6)
    ax.legend(loc="upper left", fontsize=9)
    fig.autofmt_xdate()

    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout(rect=(0, 0, 1, 0.92))
    fig.savefig(out_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_path
