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
