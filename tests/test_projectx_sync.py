"""projectx 同步烟雾测试：用 SQLite 目标库走 建表/镜像/对账 全流程。

同步与校验代码与方言无关，目标库用 SQLite 即可覆盖全部逻辑；
MySQL 方言仅影响 DDL（见 projectx.schema._MYSQL_DDL）。
"""

from __future__ import annotations

import pandas as pd
import pytest
from sqlalchemy import create_engine, text

from fundadmin.clients import store
from fundadmin.clients.schema import init_db as init_local
from fundadmin.projectx import sync as px_sync
from fundadmin.projectx.schema import init_db as init_projectx

_PRODUCTS = ("SCD704", "SCY282")


@pytest.fixture()
def engines(monkeypatch, tmp_path):
    monkeypatch.setenv("FUND_DB_URL", f"sqlite:///{tmp_path / 'local.db'}")
    src = init_local()
    dst = create_engine(f"sqlite:///{tmp_path / 'projectx.db'}")
    init_projectx(dst)
    return src, dst


def _seed_positions(dates: list[str]) -> pd.DataFrame:
    rows = []
    for d in dates:
        for i, code in enumerate(_PRODUCTS):
            rows.append(
                {
                    "as_of_date": d,
                    "product_code": code,
                    "product_name": f"产品{code}",
                    "broker": "cicc" if i == 0 else "",
                    "ticker": "AAPL" if i == 0 else "",
                    "instrument_name": "苹果" if i == 0 else "",
                    "asset_class": "equity" if i == 0 else None,
                    "quantity": 100.0 + i,
                    "market_value_cny": 12345.6 + i,
                    "weight": 0.5,
                    "contract_no": "C1" if i == 0 else "",
                }
            )
    return pd.DataFrame(rows)


def _seed_valuation(dates: list[str]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "as_of_date": d,
                "product_code": code,
                "product_name": f"产品{code}",
                "unit_nav": 1.0 + i / 100,
                "asset_nav": 10000.0 + i,
                "total_market_value_cny": 9999.0 + i,
            }
            for d in dates
            for i, code in enumerate(_PRODUCTS)
        ]
    )


def test_sync_and_verify_roundtrip(engines):
    src, dst = engines
    assert store.upsert_positions(_seed_positions(["2026-08-18", "2026-08-19"])) == 4
    assert store.upsert_product_valuation(_seed_valuation(["2026-08-18", "2026-08-19"])) == 4

    stats = px_sync.sync_all(source=src, target=dst)
    assert {s["table"] for s in stats} == {"fund_positions", "fund_product_valuation"}
    assert all(s["rows"] == 4 and s["dates"] == 2 for s in stats)

    assert px_sync.verify(source=src, target=dst) == []

    # 幂等：重复同步仍一致，且不产生重复行
    px_sync.sync_all(source=src, target=dst)
    assert px_sync.verify(source=src, target=dst) == []
    with dst.connect() as conn:
        assert conn.execute(text("SELECT COUNT(*) FROM fund_positions")).scalar() == 4


def test_sync_propagates_local_changes(engines):
    src, dst = engines
    store.upsert_positions(_seed_positions(["2026-08-18", "2026-08-19"]))
    px_sync.sync_positions(source=src, target=dst)

    # 本地删除一个估值日、新增一个估值日 → 全量镜像后远端应完全跟随
    with src.begin() as conn:
        conn.execute(text("DELETE FROM fund_positions WHERE as_of_date = '2026-08-19'"))
    store.upsert_positions(_seed_positions(["2026-08-20"]))
    px_sync.sync_positions(source=src, target=dst)

    with dst.connect() as conn:
        dates = {
            str(r[0])[:10]
            for r in conn.execute(text("SELECT DISTINCT as_of_date FROM fund_positions"))
        }
    assert dates == {"2026-08-18", "2026-08-20"}
    # 本测试只涉及持仓表；净值表两侧均为空的告警属预期，不在此断言范围
    position_issues = [i for i in px_sync.verify(source=src, target=dst) if "fund_positions" in i]
    assert position_issues == []


def test_sync_with_date_filter_and_mismatch(engines):
    src, dst = engines
    store.upsert_positions(_seed_positions(["2026-08-18", "2026-08-19"]))

    stat = px_sync.sync_positions(source=src, target=dst, dates=["2026-08-18"])
    assert stat["dates"] == 1 and stat["rows"] == 2

    with dst.connect() as conn:
        assert conn.execute(text("SELECT COUNT(*) FROM fund_positions")).scalar() == 2

    # 指定日期同步不触碰其它日期 → verify 应报告 08-19 缺失
    issues = px_sync.verify(source=src, target=dst)
    assert any("2026-08-19" in i for i in issues)


def test_row_values_survive_roundtrip(engines):
    src, dst = engines
    store.upsert_positions(_seed_positions(["2026-08-18"]))
    store.upsert_product_valuation(_seed_valuation(["2026-08-18"]))
    px_sync.sync_all(source=src, target=dst)

    local = store.load_positions(as_of_date="2026-08-18", engine=src)
    with dst.connect() as conn:
        remote = pd.read_sql(
            text("SELECT * FROM fund_positions WHERE as_of_date = '2026-08-18'"), conn
        )
    key = ["product_code", "broker", "ticker", "contract_no"]
    merged = local.merge(remote, on=key, suffixes=("_l", "_r"))
    assert len(merged) == 2
    for _, row in merged.iterrows():
        assert abs(row["market_value_cny_l"] - row["market_value_cny_r"]) < 1e-9
        assert row["instrument_name_l"] == row["instrument_name_r"]
