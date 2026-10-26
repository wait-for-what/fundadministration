"""复权净值（分红再投资口径）回归测试。

守护的关键不变量：
- 标签归一化（中金模板带尾部冒号、申万模板不带）与数值解析。
- 产品识别必须排除 `种子2号` 这类"名称含子串但属不同产品"的附件。
- 累计派现缺口按前值填充，不能把"字段缺失"当成当日分红。
- 复权恒等：adj_nav == unit_nav + 累计派现；分红当日不再是负收益。
- 事件在派现序列上求差分，窗口之前就已派完的累计分红不得报成窗口内事件。

这些是资金系统的关键不变量（scripts-tests-4）。
"""

from __future__ import annotations

import pandas as pd
import pytest

from fundadmin.portfolio import adjusted_nav as an


# ---------- 标签 / 数值解析 ----------

def test_normalize_label_strips_both_colon_styles():
    assert an._normalize_label("累计派现金额：") == "累计派现金额"
    assert an._normalize_label("累计派现金额:") == "累计派现金额"
    assert an._normalize_label(" 累计派现金额 ") == "累计派现金额"


def test_to_float_rejects_placeholders():
    assert an._to_float("0.21") == 0.21
    assert an._to_float("1,234.5") == 1234.5
    assert an._to_float(3) == 3.0
    for blank in ("", "-", "--", None, "abc"):
        assert an._to_float(blank) is None


def test_parse_date_supports_both_templates():
    assert an._parse_date("估值日期：2026-10-08") == "2026-10-08"
    assert an._parse_date("日期：2026-05-21") == "2026-05-21"
    assert an._parse_date("日期:2026/5/7") == "2026-05-07"
    assert an._parse_date("没有日期") is None


# ---------- 产品识别 ----------

def test_product_from_filename_excludes_same_name_prefix_products():
    assert an._product_from_filename("弘运盛泰种子私募证券投资基金2026-10-08资产估值表.xlsx") == "SCY282"
    assert an._product_from_filename("弘运盛泰铂金8号私募证券投资基金估值表.xlsx") == "SXQ602"
    # 种子2号 是另一只基金，不能命中 种子
    assert an._product_from_filename("弘运盛泰种子2号私募证券投资基金产品 2022-01-19估值表.xls") is None
    assert an._product_from_filename("中信证券指标计算结果.xlsx") is None


# ---------- 累计派现前向填充 ----------

def _nav_frame(rows):
    return pd.DataFrame(
        rows, columns=["product_code", "product_name", "val_date", "unit_nav"]
    ).assign(val_date=lambda d: pd.to_datetime(d["val_date"]))


def test_payout_gap_is_forward_filled_not_treated_as_dividend():
    # 中间缺 累计派现金额 字段：ffill 后不应产生新的分红事件
    nav = _nav_frame([
        ("SCD704", "全球视野", "2024-01-02", 1.00),
        ("SCD704", "全球视野", "2024-01-03", 1.01),
        ("SCD704", "全球视野", "2024-01-04", 1.02),
    ])
    payout = pd.DataFrame([
        ("SCD704", "全球视野", "2024-01-02", 0.00, None, None),
        ("SCD704", "全球视野", "2024-01-04", 0.10, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    merged = an._payout_forward_filled(nav, payout)
    # 前向填充：缺口沿用上一已知值 0.00，01-04 才跳到 0.10
    assert list(merged["cum_payout"]) == [0.00, 0.00, 0.10]

    events = an.detect_dividend_events(merged[["product_code", "product_name", "val_date", "cum_payout"]])
    # 只有 01-04 这一处真实跳变，中间缺口不再产生噪声事件
    assert len(events) == 1
    assert events.iloc[0]["val_date"] == pd.Timestamp("2024-01-04")


def test_gap_without_step_produces_no_event():
    nav = _nav_frame([
        ("SCD704", "全球视野", "2024-01-02", 1.00),
        ("SCD704", "全球视野", "2024-01-03", 1.01),
        ("SCD704", "全球视野", "2024-01-04", 1.02),
    ])
    payout = pd.DataFrame([
        ("SCD704", "全球视野", "2024-01-02", 0.00, None, None),
        ("SCD704", "全球视野", "2024-01-04", 0.00, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    merged = an._payout_forward_filled(nav, payout)
    events = an.detect_dividend_events(merged[["product_code", "product_name", "val_date", "cum_payout"]])
    assert events.empty, "字段缺失不得被误判为分红"


def test_payout_forward_fill_handles_empty_payout():
    nav = _nav_frame([("SNJ280", "铂金2号", "2024-01-02", 1.00)])
    merged = an._payout_forward_filled(nav, pd.DataFrame())
    assert list(merged["cum_payout"]) == [0.0]


# ---------- 复权计算 ----------

def test_adjusted_nav_identity_and_dividend_day_is_not_a_loss():
    nav = _nav_frame([
        ("SXQ602", "铂金8号", "2026-01-14", 1.400),
        ("SXQ602", "铂金8号", "2026-01-15", 1.042),
        ("SXQ602", "铂金8号", "2026-01-16", 1.043),
    ])
    payout = pd.DataFrame([
        ("SXQ602", "铂金8号", "2026-01-14", 0.0000, None, None),
        ("SXQ602", "铂金8号", "2026-01-15", 0.3454, 1.042, 1.3874),
        ("SXQ602", "铂金8号", "2026-01-16", 0.3454, 1.043, 1.3884),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    adj = an.build_adjusted_nav(nav, payout)

    # 复权恒等
    assert abs(adj.loc[1, "adj_nav"] - (1.042 + 0.3454)) < 1e-9
    # 分红当日回报必须剔除分红本身：0.3454/1.400 = 24.7%，不是 25.6%
    assert adj.loc[1, "daily_total_return"] == pytest.approx(1.3874 / 1.400 - 1, abs=1e-9)
    assert abs(adj.loc[1, "daily_total_return"]) < 0.01
    assert adj.loc[1, "adj_nav"] > adj.loc[0, "unit_nav"] * 0.99


def test_daily_return_uses_day_dividend_not_cumulative():
    """回归：曾用累计派现代替当日分红，把历史已派金额重复计入，导致除息日收益被高估。"""
    # 当日分红 0.05，但此前已累计派过 0.10
    nav = _nav_frame([
        ("SCY282", "种子", "2022-08-02", 1.5519),
        ("SCY282", "种子", "2022-08-03", 1.5420),
    ])
    payout = pd.DataFrame([
        ("SCY282", "种子", "2022-08-02", 0.10, None, None),
        ("SCY282", "种子", "2022-08-03", 0.15, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    adj = an.build_adjusted_nav(nav, payout)
    # 正确：(1.5420 + 0.05)/1.5519 - 1；错误做法会得到 (1.5420 + 0.15)/1.5519 - 1
    assert adj.loc[1, "daily_total_return"] == pytest.approx((1.5420 + 0.05) / 1.5519 - 1, abs=1e-9)
    assert adj.loc[1, "daily_total_return"] < 0.03


# ---------- 分红再投资（份额法） ----------

def test_reinvest_shares_buy_at_ex_nav():
    """除息日按除息净值买入新份额：0.3454/1.042。"""
    nav = _nav_frame([
        ("SXQ602", "铂金8号", "2026-01-14", 1.400),
        ("SXQ602", "铂金8号", "2026-01-15", 1.042),
    ])
    payout = pd.DataFrame([
        ("SXQ602", "铂金8号", "2026-01-14", 0.0000, None, None),
        ("SXQ602", "铂金8号", "2026-01-15", 0.3454, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    rein = an.build_reinvested_nav(nav.merge(
        payout[["product_code", "val_date", "cum_payout"]],
        on=["product_code", "val_date"], how="left").assign(
        cum_payout=lambda d: d["cum_payout"].fillna(0.0))[
        ["product_code", "product_name", "val_date", "unit_nav", "cum_payout"]])

    assert rein.loc[0, "reinvest_shares"] == 1.0
    assert rein.loc[1, "dividend_on_day"] == pytest.approx(0.3454)
    assert rein.loc[1, "reinvest_shares"] == pytest.approx(1 + 0.3454 / 1.042)
    # 除息当日回报是 -0.9% 的市场波动，不是 -25.6%
    assert rein.loc[1, "day_total_return"] == pytest.approx(1.3874 / 1.400 - 1, abs=1e-9)


def test_reinvest_matches_raw_when_no_dividends():
    """交叉校验：无分红产品，分红再投资曲线必须与未复权净值完全一致。"""
    nav = _nav_frame([
        ("SNJ280", "铂金2号", "2024-01-02", 1.00),
        ("SNJ280", "铂金2号", "2024-02-01", 1.10),
        ("SNJ280", "铂金2号", "2024-03-01", 0.95),
    ])
    zero = pd.DataFrame([("SNJ280", "铂金2号", d, 0.0, None, None)
                         for d in ("2024-01-02", "2024-03-01")],
                        columns=["product_code", "product_name", "val_date", "cum_payout",
                                 "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))
    merged = an._payout_forward_filled(nav, zero)

    rein = an.build_reinvested_nav(
        merged[["product_code", "product_name", "val_date", "unit_nav", "cum_payout"]])
    assert list(rein["reinvest_shares"].round(12)) == [1.0, 1.0, 1.0]
    pd.testing.assert_series_equal(
        rein["reinvest_nav"].reset_index(drop=True),
        rein["unit_nav"].reset_index(drop=True),
        check_names=False,
    )


def test_reinvest_does_not_double_count_nav_growth():
    """回归：曾让 shares 复利增长市场回报，导致 shares*nav 重复计入净值涨幅。"""
    nav = _nav_frame([
        ("SNJ280", "铂金2号", "2024-01-02", 1.00),
        ("SNJ280", "铂金2号", "2024-02-01", 1.10),
    ])
    merged = nav.copy()
    merged["cum_payout"] = 0.0
    rein = an.build_reinvested_nav(
        merged[["product_code", "product_name", "val_date", "unit_nav", "cum_payout"]])
    # 净值涨 10%，组合市值也只能涨 10%
    assert rein["reinvest_shares"].iloc[-1] == 1.0
    assert rein["reinvest_return_index"].iloc[-1] == pytest.approx(1.10)


def test_reinvest_ignores_payout_before_window():
    """窗口之前已派完的分红不得在首日被当成当日分红再投资。

    ffill 只填补已知值之间的缺口，不会向前回溯到首个观测之前；
    因此窗口内每日都上报 0.435 时，首日也不应产生一笔当日分红。
    """
    nav = _nav_frame([
        ("SLL384", "铂金1号", "2023-02-20", 0.88),
        ("SLL384", "铂金1号", "2023-06-20", 0.95),
    ])
    payout = pd.DataFrame([
        ("SLL384", "铂金1号", "2022-01-18", 0.435, None, None),
        ("SLL384", "铂金1号", "2023-02-20", 0.435, None, None),
        ("SLL384", "铂金1号", "2023-06-20", 0.435, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))
    merged = an._payout_forward_filled(nav, payout)
    assert list(merged["cum_payout"]) == [0.435, 0.435]

    rein = an.build_reinvested_nav(
        merged[["product_code", "product_name", "val_date", "unit_nav", "cum_payout"]])
    assert list(rein["dividend_on_day"]) == [0.0, 0.0]
    assert list(rein["reinvest_shares"].round(12)) == [1.0, 1.0]
    # 无窗口内分红 -> 再投资收益 == 未复权收益
    assert rein["reinvest_return_index"].iloc[-1] == pytest.approx(0.95 / 0.88)


def test_payout_ffill_does_not_look_backward():
    """派现序列首个观测早于净值窗口时，不应回填到窗口内（否则等于虚构窗口前派现）。"""
    nav = _nav_frame([("SLL384", "铂金1号", "2023-02-20", 0.88)])
    payout = pd.DataFrame([
        ("SLL384", "铂金1号", "2022-01-18", 0.435, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))
    merged = an._payout_forward_filled(nav, payout)
    assert list(merged["cum_payout"]) == [0.0]


def test_build_nav_variants_aligns_columns():
    nav = _nav_frame([
        ("SCD704", "全球视野", "2024-01-02", 1.00),
        ("SCD704", "全球视野", "2024-02-01", 1.05),
        ("SCD704", "全球视野", "2024-03-01", 1.10),
    ])
    payout = pd.DataFrame([
        ("SCD704", "全球视野", "2024-01-02", 0.00, None, None),
        ("SCD704", "全球视野", "2024-03-01", 0.05, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    out = an.build_nav_variants(nav, payout)
    for col in ("unit_nav", "cum_payout", "adj_nav", "total_return_index",
                "dividend_on_day", "reinvest_nav", "reinvest_return_index"):
        assert col in out.columns
    assert len(out) == 3
    assert out["reinvest_nav"].notna().all()


def test_total_return_index_normalised_to_one_at_start():
    nav = _nav_frame([
        ("SNJ280", "铂金2号", "2024-01-02", 1.00),
        ("SNJ280", "铂金2号", "2024-06-03", 1.20),
    ])
    adj = an.build_adjusted_nav(nav, pd.DataFrame())
    assert abs(adj["total_return_index"].iloc[0] - 1.0) < 1e-12
    assert abs(adj["total_return_index"].iloc[1] - 1.20) < 1e-12

# ---------- 事件窗口标注 ----------

def test_payout_before_nav_window_is_not_reported_as_in_window_event():
    # 铂金1号 的 0.435 在 2022-01-18 已派完，而净值序列从 2023-02-20 才开始
    payout = pd.DataFrame([
        ("SLL384", "铂金1号", "2022-01-17", 0.000, None, None),
        ("SLL384", "铂金1号", "2022-01-18", 0.435, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))
    nav = _nav_frame([
        ("SLL384", "铂金1号", "2023-02-20", 0.880),
        ("SLL384", "铂金1号", "2023-02-21", 0.886),
    ])

    events = an.detect_dividend_events(payout, nav_window=nav)
    assert len(events) == 1
    assert events.iloc[0]["val_date"] == pd.Timestamp("2022-01-18")
    assert events.iloc[0]["amount_per_unit"] == 0.435
    assert not bool(events.iloc[0]["in_nav_window"])
    # 该分红只对复权水平做常数平移，不影响收益率
    adj = an.build_adjusted_nav(nav, payout)
    assert abs(adj["total_return_index"].iloc[-1] - (0.886 / 0.880)) < 1e-12


def test_dividend_is_corroborated_by_nav_drop():
    """分红当日净值应同步下跌约 amount/nav_before，否则该事件可疑。"""
    # 真实形态：0.10 分红叠加约 -1% 波动 -> 净值 1.00 -> 0.895
    nav = _nav_frame([
        ("SCY282", "种子", "2022-07-04", 1.000),
        ("SCY282", "种子", "2022-07-05", 0.895),
    ])
    payout = pd.DataFrame([
        ("SCY282", "种子", "2022-07-04", 0.00, None, None),
        ("SCY282", "种子", "2022-07-05", 0.10, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    events = an.detect_dividend_events(payout, nav_window=nav)
    assert len(events) == 1
    assert bool(events.iloc[0]["explained_by_payout"])
    assert events.iloc[0]["nav_change"] == pytest.approx(-0.105, abs=1e-9)


def test_first_observation_event_is_flagged_when_nav_does_not_drop():
    """派现序列从非零值开始时首行会被记为事件；净值没跌则不应视为已解释。"""
    nav = _nav_frame([
        ("SCY282", "种子", "2022-07-04", 1.000),
        ("SCY282", "种子", "2022-07-05", 1.000),
    ])
    payout = pd.DataFrame([
        ("SCY282", "种子", "2022-07-05", 0.10, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    events = an.detect_dividend_events(payout, nav_window=nav)
    assert len(events) == 1
    assert not bool(events.iloc[0]["explained_by_payout"]), "净值未同步下跌，应判为不可解释"


def test_multiple_dividends_accumulate_on_payout_series():
    payout = pd.DataFrame([
        ("SCD704", "全球视野", "2020-11-23", 0.000, None, None),
        ("SCD704", "全球视野", "2020-11-24", 0.045, None, None),
        ("SCD704", "全球视野", "2021-01-27", 0.045, None, None),
        ("SCD704", "全球视野", "2021-01-28", 0.140, None, None),
    ], columns=["product_code", "product_name", "val_date", "cum_payout",
                "table_unit_nav", "table_cum_unit_nav"]).assign(
        val_date=lambda d: pd.to_datetime(d["val_date"]))

    events = an.detect_dividend_events(payout)
    # 同日重复上报同一累计值不算两次分红
    assert len(events) == 2
    assert events["amount_per_unit"].tolist() == pytest.approx([0.045, 0.095])
