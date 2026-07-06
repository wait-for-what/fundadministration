"""代码评审整改的回归测试（Phase 1-2）。

覆盖：HTML 转义、CSV 编码兜底、表头识别失败即报错、客户邮件失败/跳过区分与
客户隔离、客户单位净值完整性门槛、SQLite WAL。这些是资金系统的关键不变量，
历史上无测试守护（scripts-tests-4）。
"""

from __future__ import annotations

import warnings
from datetime import date
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import pytest


# ---------- HTML 转义（notifications-5 / security-3 / email-ingest-7） ----------

def test_client_nav_email_escapes_injection():
    from fundadmin.portfolio.client_notifier import _build_client_html

    holdings = [{"product_name": "<script>x</script>", "unit_nav": 1.0, "shares": 10.0}]
    out = _build_client_html("<img src=x onerror=alert(1)>", holdings, date(2026, 6, 25))
    assert "<script>x</script>" not in out
    assert "<img src=x onerror" not in out
    assert "&lt;script&gt;" in out and "&lt;img" in out


def test_matrix_table_escapes_broker_ticker():
    from fundadmin.portfolio import notifier

    df = pd.DataFrame([{"group_key": "<b>E</b>", "company": "C", "ticker": "<b>E</b>", "weight": 0.5}])
    res = [{
        "product_name": "<u>P</u>", "unit_nav": 1.0, "asset_nav": 1.0, "nav": 1.0,
        "total_holdings": 1, "total_market_value_cny": 100.0, "holdings_raw": df,
    }]
    t = notifier.build_weight_matrix_table(res, date(2026, 6, 25))
    assert "<b>E</b>" not in t and "<u>P</u>" not in t
    assert "&lt;b&gt;E" in t


def test_by_broker_card_escapes_security_name():
    from fundadmin.portfolio import by_broker_email

    bb = pd.DataFrame([{
        "broker": "cicc", "company": "<svg/onload=1>", "ticker": "<svg/onload=1>",
        "shares": 100, "market_value_cny": 100.0, "weight": 0.5, "group_key": "g",
    }])
    card, _ = by_broker_email._render_product_card(
        product_name="<x>", product_code="P1", as_of="2026-06-25", by_broker=bb
    )
    assert "<svg/onload=1>" not in card and "&lt;svg" in card


# ---------- CSV 编码兜底 + 表头识别（parsers-3 / parsers-8） ----------

def test_read_csv_robust_decodes_gbk(tmp_path: Path):
    from fundadmin.portfolio.parsers.common import read_csv_robust

    p = tmp_path / "gbk.csv"
    p.write_text("标的名称,市值（人民币）\n中国平安,1000\n", encoding="gbk")
    df = read_csv_robust(p, header=None)
    assert "标的名称" in df.iloc[0].tolist()


def test_cicc_header_detect_returns_minus_one_on_miss():
    from fundadmin.portfolio.parsers.cicc import _detect_header_row

    df = pd.DataFrame([["foo", "bar"], ["1", "2"]])
    assert _detect_header_row(df) == -1


def test_cicc_holdings_raises_when_header_not_found(tmp_path: Path):
    from fundadmin.portfolio.parsers.cicc import parse_cicc_holdings

    p = tmp_path / "bad.csv"
    p.write_text("foo,bar\n1,2\n", encoding="utf-8-sig")
    with pytest.raises(ValueError, match="未能定位表头"):
        parse_cicc_holdings(p)


# ---------- 客户邮件：失败 vs 跳过、客户隔离（notifications-4 / clients-store-1 隔离） ----------

def _clients_df():
    return pd.DataFrame([
        {"custname": "甲", "prodcode": "SXQ602", "product_name": "铂金8号", "email": "a@x.com", "持有份额": 100.0},
        {"custname": "乙", "prodcode": "SXQ602", "product_name": "铂金8号", "email": "b@x.com", "持有份额": 200.0},
        {"custname": "丙", "prodcode": "SXQ602", "product_name": "铂金8号", "email": "", "持有份额": 50.0},
    ])


def test_client_notify_distinguishes_failed_from_skipped():
    from fundadmin.portfolio import client_notifier as cn

    navmap = {"铂金8号": {"unit_nav": 1.5, "asset_nav": 1.6}}

    def fake_send(cfg, *, subject, html_body, to_addrs, **kw):
        if to_addrs == ["b@x.com"]:
            raise RuntimeError("smtp boom")

    with patch.object(cn, "load_clients", return_value=_clients_df()), \
         patch.object(cn, "load_nav_map", return_value=navmap), \
         patch.object(cn, "send_html_email", side_effect=fake_send):
        warnings.simplefilter("ignore")
        s = cn.send_client_nav_emails(date(2026, 6, 25), Path("/tmp/x.xlsx"), object())

    assert s["sent"] == 1 and s["failed"] == 1 and s["skipped"] == 1 and s["eligible"] == 2
    assert s["failed_clients"] == ["乙"]


def test_client_isolation_each_email_only_own_holdings():
    """每封客户邮件只能包含该客户自己的持仓（client isolation 不变量）。"""
    from fundadmin.portfolio import client_notifier as cn

    navmap = {"铂金8号": {"unit_nav": 1.5, "asset_nav": 1.6}}
    captured: dict[str, str] = {}

    def capture(cfg, *, subject, html_body, to_addrs, **kw):
        captured[to_addrs[0]] = html_body

    with patch.object(cn, "load_clients", return_value=_clients_df()), \
         patch.object(cn, "load_nav_map", return_value=navmap), \
         patch.object(cn, "send_html_email", side_effect=capture):
        cn.send_client_nav_emails(date(2026, 6, 25), Path("/tmp/x.xlsx"), object())

    assert set(captured) == {"a@x.com", "b@x.com"}
    assert "甲" in captured["a@x.com"] and "乙" not in captured["a@x.com"]
    assert "乙" in captured["b@x.com"] and "甲" not in captured["b@x.com"]


# ---------- 客户单位净值门槛（notifications-2） ----------

def test_client_unit_nav_gate_flags_missing():
    from fundadmin.portfolio.cross_broker_report import PRODUCT_CONFIG
    from fundadmin.portfolio.reports import client_unit_nav_missing

    names = [c["name"] for c in PRODUCT_CONFIG]
    results = [
        {"product_name": names[0], "unit_nav": 1.2, "asset_nav": 1.3},
        {"product_name": names[1], "unit_nav": None, "asset_nav": 1.4},
    ]
    miss = client_unit_nav_missing(results)
    assert names[1] in miss and names[0] not in miss


# ---------- 事务原子性（clients-store-3 / ops-sync-11 / raw-ingest gap） ----------

def _one_position() -> pd.DataFrame:
    return pd.DataFrame([{
        "as_of_date": "2026-06-02", "product_code": "SXQ602", "product_name": "铂金8号",
        "broker": "cicc", "instrument_name": "TSLA", "ticker": "TSLA",
        "quantity": 100, "market_value_cny": 200.0,
    }])


def test_replace_positions_rolls_back_on_failure(monkeypatch, tmp_path):
    from fundadmin.clients import store
    from fundadmin.clients.schema import init_db

    monkeypatch.setenv("FUND_DB_URL", f"sqlite:///{tmp_path / 'rp.db'}")
    init_db()
    store.upsert_positions(_one_position())
    assert len(store.load_positions(product_code="SXQ602")) == 1

    def boom(*a, **k):
        raise RuntimeError("upsert boom")

    monkeypatch.setattr(store, "_upsert_conn", boom)
    with pytest.raises(RuntimeError):
        store.replace_positions("SXQ602", "2026-06-02", _one_position())

    # 删除应随事务回滚：旧快照仍在，绝不会出现空快照窗口。
    assert len(store.load_positions(product_code="SXQ602")) == 1


def test_insert_attachment_with_rows_is_atomic(monkeypatch, tmp_path):
    from sqlalchemy import text

    from fundadmin.clients import store
    from fundadmin.clients.schema import init_db

    monkeypatch.setenv("FUND_DB_URL", f"sqlite:///{tmp_path / 'raw.db'}")
    init_db()
    meta = {
        "sha256": "abc123", "file_name": "x.csv", "file_suffix": ".csv", "broker": "cicc",
        "product_code": "SXQ602", "as_of_date": "2026-06-02", "attachment_type": "holdings",
        "sheet_count": 1, "row_count": 1, "inbox_dir": "/tmp",
    }

    def boom(_ingest_id):
        raise RuntimeError("rows build failed")

    with pytest.raises(RuntimeError):
        store.insert_attachment_with_rows(meta, boom)

    # 行写失败 -> 整个事务回滚 -> sha256 未登记 -> 重投递仍可入库（无孤儿附件）。
    eng = store.get_engine()
    with eng.connect() as c:
        assert c.execute(text("SELECT COUNT(*) FROM attachment_ingest WHERE sha256='abc123'")).scalar() == 0

    def good(ingest_id):
        return pd.DataFrame([{
            "ingest_id": ingest_id, "sheet_index": 0, "sheet_name": "csv",
            "row_index": 0, "cells_json": "[]",
        }])

    iid, is_new, n_rows = store.insert_attachment_with_rows(meta, good)
    assert is_new and n_rows == 1
    assert len(store.load_raw_sheet_rows(iid)) == 1
    # 同 sha256 二次入库：is_new=False，幂等。
    iid2, is_new2, _ = store.insert_attachment_with_rows(meta, good)
    assert iid2 == iid and not is_new2


# ---------- JSON 状态原子写 / 损坏恢复（ops-sync-3 / reliability-2） ----------

def test_atomic_write_roundtrip(tmp_path):
    from fundadmin.portfolio import operations as ops

    p = tmp_path / "sub" / "state.json"
    ops._atomic_write_text(p, '{"a": 1}')
    assert p.read_text(encoding="utf-8") == '{"a": 1}'


def test_email_sync_state_recovers_from_corruption(tmp_path):
    from fundadmin.portfolio import operations as ops

    p = tmp_path / "email_sync_state.json"
    p.write_text("{ this is not valid json", encoding="utf-8")
    st = ops._load_email_sync_state(p)
    assert st == {"version": 1, "processed_messages": {}}


def test_publish_state_raises_on_corruption(tmp_path):
    from fundadmin.portfolio import operations as ops

    # 发布状态损坏绝不静默重置（否则会向全体客户重发）：必须报错由人工介入。
    p = tmp_path / "publish_state.json"
    p.write_text("{ not json", encoding="utf-8")
    with pytest.raises(Exception):
        ops._load_publish_state(p)


# ---------- 数值解析正确性（parsers-1 / parsers-5 / parsers-2 CRITICAL） ----------

def test_to_float_parentheses_negative():
    from fundadmin.portfolio.parsers.common import to_float

    assert to_float("(1,234.50)") == -1234.5
    assert to_float("(500)") == -500.0
    assert to_float("1,234.50") == 1234.5
    assert to_float("100.00%") == 1.0  # 百分比行为不变


def test_to_int_rounds_not_truncates():
    from fundadmin.portfolio.parsers.common import to_int

    assert to_int("999.9999999") == 1000   # 浮点噪声不再少算一股
    assert to_int(1000.0000001) == 1000
    assert to_int("1000.4") == 1000
    assert to_int(None) is None


def test_nav_fallback_skips_percent_column(tmp_path):
    """CRITICAL parsers-2：兜底取数必须跳过 '市值占净值%' 列，不能把 1.0 当 NAV。"""
    from fundadmin.portfolio.parsers.valuation_nav import parse_nav_from_valuation

    p = tmp_path / "val.csv"
    # 资产净值行无已知 NAV 列头 -> 触发兜底；最右是 100.00%（占比），真实 NAV 在其左侧。
    p.write_text("科目,市值,市值占净值%\n资产净值,12345678.90,100.00%\n", encoding="utf-8-sig")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        nav = parse_nav_from_valuation(p)
    assert nav == 12345678.90


# ---------- 客户身份歧义即跳过（clients-store-1 / notifications-3） ----------

def test_same_name_two_emails_is_ambiguous_not_sent():
    from fundadmin.portfolio import client_notifier as cn

    clients = pd.DataFrame([
        {"custname": "张伟", "prodcode": "SXQ602", "product_name": "铂金8号", "email": "a@x.com", "持有份额": 100.0},
        {"custname": "张伟", "prodcode": "SCY282", "product_name": "种子", "email": "b@y.com", "持有份额": 200.0},
    ])
    navmap = {"铂金8号": {"unit_nav": 1.5, "asset_nav": 1.6}, "种子": {"unit_nav": 2.0, "asset_nav": 2.1}}
    sent_to: list[str] = []

    with patch.object(cn, "load_clients", return_value=clients), \
         patch.object(cn, "load_nav_map", return_value=navmap), \
         patch.object(cn, "send_html_email", side_effect=lambda *a, **k: sent_to.append(k["to_addrs"][0])):
        warnings.simplefilter("ignore")
        s = cn.send_client_nav_emails(date(2026, 6, 25), Path("/tmp/x.xlsx"), object())

    # 同名两邮箱 -> 身份歧义 -> 不发，避免把甲的持仓发给乙。
    assert s["ambiguous"] == 1 and s["sent"] == 0
    assert sent_to == []


# ---------- FX 币种精确配对（FX gap） ----------

def test_currency_token_pairing():
    from fundadmin.portfolio.cross_broker_report import _currency_token

    assert _currency_token("CITIC_USD_Underlying_2026-06-02") == "USD"
    assert _currency_token("CITIC_HKD_Balance_2026-06-02") == "HKD"
    assert _currency_token("some_file_no_ccy") is None


# ---------- 发件人白名单只匹配地址、不匹配显示名（email-ingest-1） ----------

def test_sender_allowed_rejects_display_name_spoof():
    from fundadmin.portfolio.operations import _sender_allowed

    allow = {"citicsec.com"}
    # 真实券商域名通过
    assert _sender_allowed("CITIC <noreply@citicsec.com>", allow) is True
    # 把可信关键词塞进显示名、真实地址是攻击者域名 -> 拒绝
    assert _sender_allowed("citicsec.com 对账 <attacker@evil.com>", allow) is False
    # 空白名单 -> 不过滤
    assert _sender_allowed("anyone@anywhere.com", set()) is True


# ---------- 解析器按表头定位列（parsers-4 / parsers-7） ----------

def test_cicc_val_mv_col_detected_by_header():
    from fundadmin.portfolio.parsers.cicc import _detect_cicc_val_mv_col

    df = pd.DataFrame([["科目", "数量", "成本", "市值（本币）", "占净值%"]])
    assert _detect_cicc_val_mv_col(df) == 3  # 按表头定位，而非写死的 11


def test_citic_deriv_col_resolved_by_header_else_position():
    from fundadmin.portfolio.parsers.citic import _resolve_citic_deriv_col

    df = pd.DataFrame(columns=["标的名称", "标的代码", "x", "y", "合约数量", "市值(人民币)"])
    assert _resolve_citic_deriv_col(df, {"市值(人民币)"}, -1) == "市值(人民币)"
    # 无匹配表头时回退到位置索引
    assert _resolve_citic_deriv_col(df, {"不存在的列"}, 1) == "标的代码"


# ---------- SQLite WAL（reliability-3） ----------

def test_sqlite_engine_enables_wal(tmp_path: Path):
    from sqlalchemy import text

    from fundadmin.db.engine import get_engine

    eng = get_engine(f"sqlite:///{tmp_path/'wal.db'}")
    with eng.connect() as c:
        assert c.execute(text("PRAGMA journal_mode")).scalar().lower() == "wal"
        assert int(c.execute(text("PRAGMA busy_timeout")).scalar()) == 30000
