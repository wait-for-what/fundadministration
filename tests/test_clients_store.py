"""clients schema / store 烟雾测试：建表 + 客户/净值/持仓 upsert 与读取。"""

from __future__ import annotations

import pandas as pd

from fundadmin.clients import store
from fundadmin.clients.schema import init_db


def test_client_and_nav_roundtrip(monkeypatch, tmp_path):
    monkeypatch.setenv("FUND_DB_URL", f"sqlite:///{tmp_path / 'fundadmin_test.db'}")
    init_db()

    clients = pd.DataFrame(
        [
            {
                "custname": "张三",
                "prodcode": "SCD704",
                "prodname": "全球视野",
                "holding_shares": 100.0,
                "email": "a@b.com",
                "mobile": "13800000000",
            }
        ]
    )
    assert store.upsert_clients(clients) == 1

    nav = pd.DataFrame(
        [{"prodcode": "SCD704", "as_of_date": "2026-05-29", "nav_unit": 1.2345, "nav_cum": 1.5}]
    )
    assert store.upsert_nav(nav) == 1

    loaded = store.load_clients()
    assert len(loaded) == 1
    assert loaded.iloc[0]["custname"] == "张三"

    latest_nav = store.load_latest_nav()
    assert len(latest_nav) == 1
    assert abs(float(latest_nav.iloc[0]["nav_unit"]) - 1.2345) < 1e-9
