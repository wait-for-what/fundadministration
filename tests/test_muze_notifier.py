"""沐泽1号每日持仓变动邮件测试。"""

from __future__ import annotations

import json
from unittest.mock import patch

import pandas as pd

from fundadmin.portfolio import muze_notifier as mn


def _holdings() -> pd.DataFrame:
    rows = [
        {
            "as_of_date": "2026-10-16",
            "ticker": "600988",
            "instrument_name": "赤峰黄金",
            "market": "上交所",
            "quantity": 260000.0,
            "price": 41.38,
            "market_value": 10758800.0,
            "unrealized_pnl": 100.0,
            "cost": 10700000.0,
        },
        {
            "as_of_date": "2026-10-19",
            "ticker": "600988",
            "instrument_name": "赤峰黄金",
            "market": "上交所",
            "quantity": 390000.0,
            "price": 41.38,
            "market_value": 16138200.0,
            "unrealized_pnl": -100.0,
            "cost": 16000000.0,
        },
    ]
    return pd.DataFrame(rows)


def _nav() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "as_of_date": "2026-10-16",
                "unit_nav": 1.373,
                "fund_asset_nav": 221832506.0,
                "stock_ratio_pct": 17.5,
                "cash_like_ratio_pct": 80.0,
            },
            {
                "as_of_date": "2026-10-19",
                "unit_nav": 1.374,
                "fund_asset_nav": 221845492.0,
                "stock_ratio_pct": 20.1,
                "cash_like_ratio_pct": 77.5,
            },
        ]
    )


def test_build_html_contains_changes():
    from fundadmin.portfolio.parsers.muze_valuation import holdings_changes

    holdings = _holdings()
    changes = holdings_changes(holdings)
    day = changes[changes["as_of_date"] == "2026-10-19"]
    html = mn.build_muze_daily_change_html(
        day_changes=day,
        nav_row=_nav().iloc[-1].to_dict(),
        holdings_today=holdings[holdings["as_of_date"] == "2026-10-19"],
        trade_date="2026-10-19",
    )
    assert "沐泽1号" in html and "2026-10-19" in html
    assert "赤峰黄金" in html and "600988" in html
    assert "加仓" in html
    assert "130,000" in html


def test_build_html_escapes_instrument_name():
    changes = pd.DataFrame(
        [
            {
                "as_of_date": "2026-10-19",
                "prev_date": "2026-10-16",
                "ticker": "000001",
                "instrument_name": "<script>alert(1)</script>",
                "market": "深交所",
                "quantity": 100.0,
                "prev_qty": 0.0,
                "qty_chg": 100.0,
                "price": 10.0,
                "chg_amount": 1000.0,
                "market_value": 1000.0,
                "action": "新进",
            }
        ]
    )
    nav_row = _nav().iloc[-1].to_dict()
    html = mn.build_muze_daily_change_html(
        day_changes=changes,
        nav_row=nav_row,
        holdings_today=_holdings().iloc[[1]],
        trade_date="2026-10-19",
    )
    assert "<script>" not in html
    assert "&lt;script&gt;" in html


def test_send_sends_to_muze_group(tmp_path):
    captured = {}

    def fake_send(cfg, *, subject, html_body, to_addrs, **kw):
        captured.update(subject=subject, html=html_body, to_addrs=to_addrs)

    with patch.object(mn, "parse_muze_dir", return_value=(_holdings(), _nav())), \
         patch.object(mn, "send_html_email", side_effect=fake_send), \
         patch.object(mn, "resolve_email_group", return_value=["a@x.com"]) as rg:
        result = mn.send_muze_daily_change_email(
            inbox_dir="/tmp/whatever",
            state_file=tmp_path / "state.json",
            smtp_config=object(),
        )

    assert result["sent"] is True
    assert result["trade_date"] == "2026-10-19"
    assert captured["to_addrs"] == ["a@x.com"]
    assert "赤峰黄金" in captured["html"]
    rg.assert_called_with("muze")
    state = json.loads((tmp_path / "state.json").read_text(encoding="utf-8"))
    assert state["last_sent_date"] == "2026-10-19"


def test_send_skips_when_no_changes(tmp_path):
    with patch.object(mn, "parse_muze_dir", return_value=(_holdings(), _nav())), \
         patch.object(mn, "send_html_email") as send_mock:
        result = mn.send_muze_daily_change_email(
            inbox_dir="/tmp/whatever",
            trade_date="2026-10-16",
            state_file=tmp_path / "state.json",
            smtp_config=object(),
        )

    assert result["sent"] is False
    assert result["reason"] == "no_changes"
    send_mock.assert_not_called()


def test_send_skips_already_sent(tmp_path):
    state_file = tmp_path / "state.json"
    state_file.write_text(json.dumps({"last_sent_date": "2026-10-19"}), encoding="utf-8")

    with patch.object(mn, "parse_muze_dir", return_value=(_holdings(), _nav())), \
         patch.object(mn, "send_html_email") as send_mock:
        result = mn.send_muze_daily_change_email(
            inbox_dir="/tmp/whatever",
            state_file=state_file,
            smtp_config=object(),
        )

    assert result["sent"] is False
    assert result["reason"] == "already_sent"
    send_mock.assert_not_called()


def test_dry_run_does_not_send(tmp_path):
    with patch.object(mn, "parse_muze_dir", return_value=(_holdings(), _nav())), \
         patch.object(mn, "send_html_email") as send_mock, \
         patch.object(mn, "resolve_email_group", return_value=["a@x.com"]):
        result = mn.send_muze_daily_change_email(
            inbox_dir="/tmp/whatever",
            state_file=tmp_path / "state.json",
            dry_run=True,
        )

    assert result["sent"] is False
    assert result["reason"] == "dry_run"
    assert "赤峰黄金" in result["html"]
    send_mock.assert_not_called()


def test_resolve_email_group_default_falls_back_to_email_to():
    from fundadmin.notifications import recipients

    def fake_get_env(name, default=None, required=False):
        return {"EMAIL_GROUP_DEFAULT": "", "EMAIL_TO": "a@x.com, b@x.com"}.get(name, default)

    with patch.object(recipients, "get_env", side_effect=fake_get_env):
        assert recipients.resolve_email_group("default") == ["a@x.com", "b@x.com"]


def test_resolve_email_group_muze_does_not_fall_back_to_email_to():
    from fundadmin.notifications import recipients

    def fake_get_env(name, default=None, required=False):
        return {"EMAIL_GROUP_MUZE": "", "EMAIL_TO": "a@x.com"}.get(name, default)

    with patch.object(recipients, "get_env", side_effect=fake_get_env):
        assert recipients.resolve_email_group("muze") == []


def test_send_skips_when_muze_group_empty(tmp_path):
    with patch.object(mn, "parse_muze_dir", return_value=(_holdings(), _nav())), \
         patch.object(mn, "send_html_email") as send_mock, \
         patch.object(mn, "resolve_email_group", return_value=[]):
        result = mn.send_muze_daily_change_email(
            inbox_dir="/tmp/whatever",
            state_file=tmp_path / "state.json",
            smtp_config=object(),
        )

    assert result["sent"] is False
    assert result["reason"] == "no_recipients_for_group_muze"
    send_mock.assert_not_called()


def test_resolve_email_group_explicit_wins():
    from fundadmin.notifications import recipients

    def fake_get_env(name, default=None, required=False):
        return {
            "EMAIL_GROUP_DEFAULT": "c@x.com",
            "EMAIL_TO": "a@x.com",
        }.get(name, default)

    with patch.object(recipients, "get_env", side_effect=fake_get_env):
        assert recipients.resolve_email_group("default") == ["c@x.com"]
