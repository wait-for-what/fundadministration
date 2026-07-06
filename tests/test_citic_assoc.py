"""中信收益互换协会版本估值报告解析器单元测试。

用合成 workbook 复刻协会版本「存续标的汇总」结构（标题行 + 表头行 + 明细），
覆盖：按标的代码跨合约批次汇总、市值=标的市值×参考汇率、多币种、已平仓占位过滤、
sheet 结构识别。不依赖含 PII 的真实邮件附件。
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from fundadmin.portfolio.parsers.citic_assoc import (
    HOLDINGS_SHEET,
    SUMMARY_SHEET,
    is_citic_assoc_report,
    parse_citic_assoc_holdings,
)

# 「存续标的汇总」列顺序（与真实文件一致；第 0 行为标题，第 1 行为表头）。
_COLS = [
    "合约编号\n（选填）", "交易所", "标的代码", "标的名称", "标的小类", "标的方向（选填）",
    "估值日期", "是否跨境标的", "名义数量", "计价货币", "参考汇率", "乘数（选填）",
    "标的名义本金", "标的占合约名义本金", "占用保证金", "期初价格", "估值价格",
    "标的市值", "浮动收益", "浮动盈亏", "未到账股息/债息等金额（选填）",
]


def _row(code, name, ccy, direction, qty, fx, mv, notional):
    return {
        "合约编号\n（选填）": f"{code}--x", "交易所": "XNAS", "标的代码": code,
        "标的名称": name, "标的小类": "美股", "标的方向（选填）": direction,
        "估值日期": "2026-06-25", "是否跨境标的": "是", "名义数量": qty,
        "计价货币": ccy, "参考汇率": fx, "乘数（选填）": 1, "标的名义本金": notional,
        "标的占合约名义本金": notional, "占用保证金": 0, "期初价格": 0, "估值价格": 0,
        "标的市值": mv, "浮动收益": 0, "浮动盈亏": 0, "未到账股息/债息等金额（选填）": 0,
    }


def _write_assoc_workbook(path: Path, detail_rows: list[dict]) -> None:
    # 存续标的汇总：标题行 + 表头行 + 明细
    title = pd.DataFrame([["存续标的汇总"] + [None] * (len(_COLS) - 1)], columns=_COLS)
    header = pd.DataFrame([_COLS], columns=_COLS)
    body = pd.DataFrame(detail_rows, columns=_COLS)
    holdings = pd.concat([title, header, body], ignore_index=True)
    summary = pd.DataFrame([["收益互换估值报告"], ["客户估值汇总"]])
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name=SUMMARY_SHEET, header=False, index=False)
        holdings.to_excel(w, sheet_name=HOLDINGS_SHEET, header=False, index=False)


def test_parse_assoc_aggregates_and_converts_fx(tmp_path):
    p = tmp_path / "assoc.xlsx"
    _write_assoc_workbook(
        p,
        [
            # AMZN 拆两个合约批次 → 应合并为一行
            _row("AMZN.OQ", "亚马逊", "USD", "多", 200, 6.7993, 45000.0, 44000.0),
            _row("AMZN.OQ", "亚马逊", "USD", "多", 143, 6.7993, 32864.43, 31000.0),
            # HKD 标的
            _row("0286.HK", "爱帝宫", "HKD", "多", 18466, 0.86724, 24190.46, 24238.84),
            # 已平仓占位：名义数量与市值均为 0 → 应被过滤
            _row("MSFT.OQ", "微软", "USD", "多", 0, 6.7993, 0.0, 0.0),
        ],
    )
    h = parse_citic_assoc_holdings(p)

    assert set(h["ticker"]) == {"AMZN", "0286"}  # MSFT 占位行被过滤
    amzn = h[h["ticker"] == "AMZN"].iloc[0]
    assert amzn["shares"] == 343  # 200 + 143 合并
    assert amzn["market_value_cny"] == pytest_approx((45000.0 + 32864.43) * 6.7993)
    assert amzn["cost_ccy"] == "USD"

    hkd = h[h["ticker"] == "0286"].iloc[0]
    assert hkd["market_value_cny"] == pytest_approx(24190.46 * 0.86724)
    assert hkd["cost_ccy"] == "HKD"


def test_parse_assoc_short_direction_negates(tmp_path):
    p = tmp_path / "assoc_short.xlsx"
    _write_assoc_workbook(
        p, [_row("TSLA.OQ", "特斯拉", "USD", "空", 10, 7.0, 1000.0, 900.0)]
    )
    h = parse_citic_assoc_holdings(p)
    row = h.iloc[0]
    assert row["shares"] == -10
    assert row["market_value_cny"] == pytest_approx(-1000.0 * 7.0)


def test_is_citic_assoc_report_detects_by_sheets(tmp_path):
    p = tmp_path / "assoc.xlsx"
    _write_assoc_workbook(p, [_row("AMZN.OQ", "亚马逊", "USD", "多", 1, 7.0, 7.0, 7.0)])
    assert is_citic_assoc_report(p) is True

    # 非协会版本：普通单 sheet 估值表不应被识别
    other = tmp_path / "plain.xlsx"
    pd.DataFrame({"科目代码": [1], "科目名称": ["x"]}).to_excel(other, index=False)
    assert is_citic_assoc_report(other) is False


def pytest_approx(value):
    import pytest

    return pytest.approx(value, rel=1e-9)
