"""成交流水 upsert/去重 + 持仓成本列 + 交叉核对（reconcile）单元测试。"""

from __future__ import annotations

import pandas as pd

from fundadmin.clients import store
from fundadmin.clients.schema import init_db
from fundadmin.portfolio.reconcile import (
    reconcile_holding_deltas,
    reconcile_positions_vs_trades,
)


def _setup_db(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("FUND_DB_URL", f"sqlite:///{tmp_path / 'tx_test.db'}")
    init_db()


def test_positions_carry_cost_columns(monkeypatch, tmp_path):
    _setup_db(monkeypatch, tmp_path)
    pos = pd.DataFrame(
        [
            {
                "as_of_date": "2026-06-02",
                "product_code": "SXQ602",
                "product_name": "铂金8号",
                "broker": "citic",
                "ticker": "AMZN",
                "instrument_name": "AMAZON",
                "quantity": 100,
                "market_value_cny": 1_000_000.0,
                "cost_price_local": 210.5,
                "cost_value_local": 21050.0,
                "cost_ccy": "USD",
                "contract_no": "",
            }
        ]
    )
    assert store.upsert_positions(pos) == 1
    loaded = store.load_positions(product_code="SXQ602")
    assert len(loaded) == 1
    row = loaded.iloc[0]
    assert abs(float(row["cost_price_local"]) - 210.5) < 1e-9
    assert abs(float(row["cost_value_local"]) - 21050.0) < 1e-9
    assert row["cost_ccy"] == "USD"


def test_transactions_dedup_and_occ(monkeypatch, tmp_path):
    _setup_db(monkeypatch, tmp_path)
    # 两笔同键真实重复 + 一笔不同
    tx = pd.DataFrame(
        [
            {"trade_date": "2026-06-02", "product_code": "SXQ602", "broker": "citic",
             "ticker": "CEG", "direction": "B", "open_close": "", "price_local": 250.0,
             "quantity": 10.0, "contract_no": "C1"},
            {"trade_date": "2026-06-02", "product_code": "SXQ602", "broker": "citic",
             "ticker": "CEG", "direction": "B", "open_close": "", "price_local": 250.0,
             "quantity": 10.0, "contract_no": "C1"},  # 真实重复 -> occ=1
            {"trade_date": "2026-06-02", "product_code": "SXQ602", "broker": "citic",
             "ticker": "AMZN", "direction": "S", "open_close": "", "price_local": 200.0,
             "quantity": -5.0, "contract_no": "C1"},
        ]
    )
    n = store.upsert_transactions(tx)
    assert n == 3
    loaded = store.load_transactions(product_code="SXQ602")
    assert len(loaded) == 3
    # occ 分配：CEG 同键两条 -> occ 0/1
    ceg = loaded[loaded["ticker"] == "CEG"].sort_values("occ")
    assert ceg["occ"].tolist() == [0, 1]

    # 重复入库（全历史重放）不应增加行数
    store.upsert_transactions(tx)
    assert len(store.load_transactions(product_code="SXQ602")) == 3


def test_reconcile_delta_matches(monkeypatch, tmp_path):
    _setup_db(monkeypatch, tmp_path)
    # 两个相邻估值日的持仓：AMZN 100->? , CEG 0->?
    positions = pd.DataFrame(
        [
            {"as_of_date": "2026-06-01", "product_code": "SXQ602", "product_name": "铂金8号",
             "broker": "citic", "ticker": "AMZN", "instrument_name": "AMZN", "quantity": 200,
             "contract_no": ""},
            {"as_of_date": "2026-06-02", "product_code": "SXQ602", "product_name": "铂金8号",
             "broker": "citic", "ticker": "AMZN", "instrument_name": "AMZN", "quantity": 34,
             "contract_no": ""},
            {"as_of_date": "2026-06-02", "product_code": "SXQ602", "product_name": "铂金8号",
             "broker": "citic", "ticker": "CEG", "instrument_name": "CEG", "quantity": 172,
             "contract_no": ""},
        ]
    )
    store.upsert_positions(positions)
    # 06-02 成交：AMZN 卖 166（带号 -166），CEG 买 172
    tx = pd.DataFrame(
        [
            {"trade_date": "2026-06-02", "product_code": "SXQ602", "broker": "citic",
             "ticker": "AMZN", "direction": "S", "open_close": "", "price_local": 200.0,
             "quantity": -166.0, "contract_no": ""},
            {"trade_date": "2026-06-02", "product_code": "SXQ602", "broker": "citic",
             "ticker": "CEG", "direction": "B", "open_close": "", "price_local": 250.0,
             "quantity": 172.0, "contract_no": ""},
        ]
    )
    store.upsert_transactions(tx)

    df = reconcile_holding_deltas("SXQ602")
    assert not df.empty
    assert set(df["status"]) == {"ok"}
    amzn = df[df["ticker"] == "AMZN"].iloc[0]
    assert abs(amzn["holding_delta"] - (-166.0)) < 1e-9
    assert abs(amzn["net_traded_qty"] - (-166.0)) < 1e-9


def test_reconcile_detects_mismatch(monkeypatch, tmp_path):
    _setup_db(monkeypatch, tmp_path)
    positions = pd.DataFrame(
        [
            {"as_of_date": "2026-06-01", "product_code": "SCD704", "product_name": "全球视野",
             "broker": "citic", "ticker": "AMZN", "instrument_name": "AMZN", "quantity": 200,
             "contract_no": ""},
            {"as_of_date": "2026-06-02", "product_code": "SCD704", "product_name": "全球视野",
             "broker": "citic", "ticker": "AMZN", "instrument_name": "AMZN", "quantity": 100,
             "contract_no": ""},
        ]
    )
    store.upsert_positions(positions)
    # 没有任何成交记录 -> 持仓减少 100 应被判为 mismatch
    df = reconcile_holding_deltas("SCD704")
    assert not df.empty
    amzn = df[df["ticker"] == "AMZN"].iloc[0]
    assert amzn["status"] == "mismatch"
    assert abs(amzn["holding_delta"] - (-100.0)) < 1e-9
    assert abs(amzn["net_traded_qty"]) < 1e-9


def test_reconcile_cumulative_runs(monkeypatch, tmp_path):
    _setup_db(monkeypatch, tmp_path)
    positions = pd.DataFrame(
        [
            {"as_of_date": "2026-06-02", "product_code": "SXQ602", "product_name": "铂金8号",
             "broker": "citic", "ticker": "CEG", "instrument_name": "CEG", "quantity": 172,
             "contract_no": ""},
        ]
    )
    store.upsert_positions(positions)
    tx = pd.DataFrame(
        [
            {"trade_date": "2026-06-02", "product_code": "SXQ602", "broker": "citic",
             "ticker": "CEG", "direction": "B", "open_close": "", "price_local": 250.0,
             "quantity": 172.0, "contract_no": ""},
        ]
    )
    store.upsert_transactions(tx)
    df = reconcile_positions_vs_trades("SXQ602")
    ceg = df[df["ticker"] == "CEG"].iloc[0]
    assert ceg["status"] == "ok"
