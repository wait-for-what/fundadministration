"""分券商持仓：提取/存储/展示视图的单元测试。"""

from __future__ import annotations

from datetime import date

import pandas as pd

from fundadmin.clients import store
from fundadmin.clients.schema import init_db


def _setup_db(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("FUND_DB_URL", f"sqlite:///{tmp_path / 'pb_test.db'}")
    init_db()


def _by_broker_frame() -> pd.DataFrame:
    """同一标的 TSLA 同时在中金、中信；AAPL 仅在中金。"""
    return pd.DataFrame(
        [
            {"group_key": "TSLA", "ticker": "TSLA", "company": "TESLA", "broker": "cicc",
             "shares": 100, "market_value_cny": 200.0, "weight": 0.20,
             "cost_price_local": None, "cost_value_local": None, "cost_ccy": None,
             "source_files": "cicc.xlsx"},
            {"group_key": "TSLA", "ticker": "TSLA", "company": "TESLA", "broker": "citic",
             "shares": 50, "market_value_cny": 100.0, "weight": 0.10,
             "cost_price_local": None, "cost_value_local": None, "cost_ccy": None,
             "source_files": "citic.xlsx"},
            {"group_key": "AAPL", "ticker": "AAPL", "company": "APPLE", "broker": "cicc",
             "shares": 30, "market_value_cny": 60.0, "weight": 0.06,
             "cost_price_local": None, "cost_value_local": None, "cost_ccy": None,
             "source_files": "cicc.xlsx"},
        ]
    )


def _results_payload() -> list[dict]:
    return [
        {
            "product_name": "铂金8号",
            "unit_nav": 1.2,
            "asset_nav": 1000.0,
            "nav": 1000.0,
            "total_holdings": 2,
            "total_market_value_cny": 360.0,
            "holdings_by_broker": _by_broker_frame(),
        }
    ]


def test_persist_curated_layer_writes_per_broker(monkeypatch, tmp_path):
    from fundadmin.portfolio.operations import _persist_curated_layer

    _setup_db(monkeypatch, tmp_path)
    _persist_curated_layer(_results_payload(), effective_trade_date=date(2026, 6, 2))

    loaded = store.load_positions(product_code="SXQ602")
    # TSLA 两行（中金/中信）+ AAPL 一行。
    assert len(loaded) == 3
    tsla = loaded[loaded["ticker"] == "TSLA"].sort_values("broker")
    assert tsla["broker"].tolist() == ["cicc", "citic"]
    assert abs(float(tsla[tsla["broker"] == "cicc"]["quantity"].iloc[0]) - 100.0) < 1e-9
    assert abs(float(tsla[tsla["broker"] == "citic"]["quantity"].iloc[0]) - 50.0) < 1e-9


def test_persist_curated_layer_idempotent_rerun(monkeypatch, tmp_path):
    from fundadmin.portfolio.operations import _persist_curated_layer

    _setup_db(monkeypatch, tmp_path)
    _persist_curated_layer(_results_payload(), effective_trade_date=date(2026, 6, 2))
    _persist_curated_layer(_results_payload(), effective_trade_date=date(2026, 6, 2))
    # 重跑不应翻倍。
    assert len(store.load_positions(product_code="SXQ602")) == 3


def test_persist_clears_legacy_collapsed_row(monkeypatch, tmp_path):
    """旧的 broker='' 折叠行应在重建时被清掉，避免与分券商行重复计数。"""
    from fundadmin.portfolio.operations import _persist_curated_layer

    _setup_db(monkeypatch, tmp_path)
    legacy = pd.DataFrame(
        [
            {"as_of_date": "2026-06-02", "product_code": "SXQ602", "product_name": "铂金8号",
             "broker": "", "ticker": "TSLA", "instrument_name": "TESLA", "quantity": 150,
             "market_value_cny": 300.0, "contract_no": ""},
        ]
    )
    store.upsert_positions(legacy)
    assert len(store.load_positions(product_code="SXQ602")) == 1

    _persist_curated_layer(_results_payload(), effective_trade_date=date(2026, 6, 2))
    loaded = store.load_positions(product_code="SXQ602")
    # 折叠行被清掉，只剩分券商行（TSLA x2 + AAPL）。
    assert "" not in set(loaded["broker"])
    assert len(loaded) == 3


def test_delete_positions(monkeypatch, tmp_path):
    _setup_db(monkeypatch, tmp_path)
    store.upsert_positions(_by_broker_frame().assign(
        as_of_date="2026-06-02", product_code="SXQ602", product_name="铂金8号",
        instrument_name="x", quantity=1, contract_no="",
    ))
    assert len(store.load_positions(product_code="SXQ602")) >= 1
    n = store.delete_positions("SXQ602", "2026-06-02")
    assert n >= 1
    assert store.load_positions(product_code="SXQ602").empty


def _holdings_raw_frame():
    """合并总表（每标的一行，供顶部持仓权重矩阵渲染）。"""
    return pd.DataFrame(
        [
            {"group_key": "TSLA", "company": "TESLA", "ticker": "TSLA",
             "market_value_cny": 300.0, "weight": 0.30, "shares": 150},
            {"group_key": "AAPL", "company": "APPLE", "ticker": "AAPL",
             "market_value_cny": 60.0, "weight": 0.06, "shares": 30},
        ]
    )


def test_build_by_broker_summary_html():
    from fundadmin.portfolio.by_broker_email import build_by_broker_summary_html

    results = [{
        "product_name": "铂金8号",
        "holdings_by_broker": _by_broker_frame(),
        "holdings_raw": _holdings_raw_frame(),
        "unit_nav": 1.2,
        "asset_nav": 1000.0,
    }]
    html = build_by_broker_summary_html(results, trade_date=date(2026, 6, 2))
    assert html is not None
    assert "铂金8号" in html and "SXQ602" in html
    assert "TSLA" in html and "AAPL" in html
    assert "中金" in html and "中信" in html
    # 跨券商高亮色块出现（TSLA 同时在中金/中信）。
    assert "#fff7e0" in html
    # 顶部权重矩阵保留（含总权重行与单位净值/资产净值）。
    assert "持仓权重矩阵" in html and "总权重" in html
    assert "单位净值" in html and "资产净值" in html

    # 无可展示数据 -> None（不发空邮件）。
    assert build_by_broker_summary_html([], trade_date=date(2026, 6, 2)) is None
    # 未配置 product_code 的产品被跳过。
    assert (
        build_by_broker_summary_html(
            [{"product_name": "沐泽1号", "holdings_by_broker": _by_broker_frame()}],
            trade_date=date(2026, 6, 2),
        )
        is None
    )


def test_build_by_broker_view_pivot():
    from fundadmin.clients.pages.portfolio_page import _build_by_broker_view

    # 模拟 fund_positions 明细行。
    df = pd.DataFrame(
        [
            {"ticker": "TSLA", "instrument_name": "TESLA", "broker": "cicc",
             "quantity": 100, "market_value_cny": 200.0, "weight": 0.20},
            {"ticker": "TSLA", "instrument_name": "TESLA", "broker": "citic",
             "quantity": 50, "market_value_cny": 100.0, "weight": 0.10},
            {"ticker": "AAPL", "instrument_name": "APPLE", "broker": "cicc",
             "quantity": 30, "market_value_cny": 60.0, "weight": 0.06},
        ]
    )
    view, broker_cols = _build_by_broker_view(df)

    assert "中金" in broker_cols and "中信" in broker_cols
    tsla = view[view["代码"] == "TSLA"].iloc[0]
    # TSLA 跨两个券商，合计股数 150，来源券商数 2。
    assert int(tsla["合计股数"]) == 150
    assert int(tsla["来源券商数"]) == 2
    assert abs(float(tsla["中金"]) - 100.0) < 1e-9
    assert abs(float(tsla["中信"]) - 50.0) < 1e-9
    # 合计权重 = 0.20 + 0.10。
    assert abs(float(tsla["权重"]) - 0.30) < 1e-9

    aapl = view[view["代码"] == "AAPL"].iloc[0]
    assert int(aapl["来源券商数"]) == 1
    assert int(aapl["中信"]) == 0
    # 按合计市值降序：TSLA(300) 在 AAPL(60) 之前。
    assert view.iloc[0]["代码"] == "TSLA"
