"""复权净值（分红再投资口径）计算。

背景
----
系统内 **没有** 结构化的基金分红表（`净值分红` / `权益登记日` / `每10份` 在
`raw_sheet_rows` 中命中为 0）。`分红` 关键词命中的 `公司行为` / `现货分红` / `CA`
工作表全部是**标的股票**的公司行为（基金从持仓收到的股息），方向与复权相反。

但券商估值表本身带复权口径字段，且恒等：

    累计单位净值 == 单位净值 + 累计派现金额

`累计派现金额` 是每份累计分红，`累计单位净值` 即累计净值（后复权净值）。
本模块从 `raw_sheet_rows` 还原 `累计派现金额`，与 `product_valuation` 的净值对齐后生成复权曲线。

支持两种估值表模板（标签差异仅在尾部冒号，数值均在第 2 列）：
- 申万宏源 `Sheet1`：`今日单位净值` / `累计派现金额` / `累计单位净值`
- 中金公司 `估值表明细表`：`基金单位净值：` / `累计派现金额：` / `累计单位净值：`

输入:
- SQLAlchemy Engine（默认 `schema.get_engine()`）。

输出:
- `load_payout_series`: 每 (产品, 估值日) 的每份累计派现。
- `detect_dividend_events`: 累计派现的增量即为单次分红。
- `build_adjusted_nav`: 复权净值 + 累计收益指数。

注意:
- 净值取自 `product_valuation`（已校验的净值口径），不取估值表里的原始净值，
  避免多版本附件 / 同名不同产品（`种子` vs `种子2号`）串扰。
"""

from __future__ import annotations

import json
import re
from collections.abc import Mapping

import pandas as pd
from sqlalchemy import text
from sqlalchemy.engine import Engine

from fundadmin.clients.config import NAME_TO_PRODCODE, PRODCODE_TO_NAME

# 估值表标签（去尾部冒号后）-> 规范字段名
_LABEL_FIELDS: Mapping[str, str] = {
    "累计派现金额": "cum_payout",
    "累计单位净值": "cum_unit_nav",
    "今日单位净值": "unit_nav",
    "基金单位净值": "unit_nav",
    "期末单位净值": "unit_nav",
}

_DATE_LABELS = ("估值日期", "日期")
_DATE_RE = re.compile(r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})")
# 文件名兜底取日期，如 资产估值表_20260521.xls
_FILE_DATE_RE = re.compile(r"(20\d{6})")
# 基金全名 -> 产品代码。要求名称后紧跟"私募"，以排除 `种子2号` 这类不同产品
# （`弘运盛泰种子私募证券投资基金` 命中，`弘运盛泰种子2号私募证券投资基金` 不命中）。
_FUND_NAME_RE = re.compile(
    "弘运盛泰(" + "|".join(sorted(NAME_TO_PRODCODE, key=len, reverse=True)) + ")私募"
)

_SQL_PAYOUT = """
SELECT i.ingest_id, i.product_code AS db_product_code, i.file_name, r.row_index, r.cells_json
FROM raw_sheet_rows r
JOIN attachment_ingest i ON i.ingest_id = r.ingest_id
WHERE {pred}
ORDER BY i.ingest_id, r.row_index
"""

_SQL_NAV = """
SELECT as_of_date, product_code, product_name, unit_nav, asset_nav
FROM product_valuation
WHERE unit_nav IS NOT NULL
ORDER BY as_of_date, product_code
"""


def _predicates(labels: tuple[str, ...]) -> str:
    return " OR ".join(f"r.cells_json LIKE '%{lab}%'" for lab in labels)


def _to_float(value: object) -> float | None:
    """把估值表里的文本数字转 float；'-' / '--' / 空视为缺失。"""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str):
        return None
    text_ = value.replace(",", "").replace("%", "").strip()
    if text_ in ("", "-", "--", "None", "nan"):
        return None
    try:
        return float(text_)
    except ValueError:
        return None


def _normalize_label(value: object) -> str:
    """去掉尾部中英文冒号与空白，兼容两套模板的标签写法。"""
    if not isinstance(value, str):
        return ""
    return value.strip().rstrip("：:").strip()


def _parse_date(value: object) -> str | None:
    match = _DATE_RE.search(str(value or ""))
    if not match:
        return None
    year, month, day = match.groups()
    return f"{year}-{int(month):02d}-{int(day):02d}"


def _date_from_filename(file_name: str) -> str | None:
    match = _FILE_DATE_RE.search(file_name or "")
    if not match:
        return None
    raw = match.group(1)
    return f"{raw[:4]}-{raw[4:6]}-{raw[6:]}"


def _product_from_filename(file_name: str) -> str | None:
    """按基金全名反查产品代码；无法唯一匹配时返回 None。"""
    match = _FUND_NAME_RE.search(file_name or "")
    if match:
        return NAME_TO_PRODCODE.get(match.group(1))
    # 退回附件上已入库的产品代码场景由调用方处理
    return None


def load_payout_series(*, engine: Engine | None = None) -> pd.DataFrame:
    """从 raw 估值表还原每 (产品, 估值日) 的每份累计派现。

    同一 (产品, 估值日) 可能有多版本附件（回补重传 / "以最新为准"），取 ingest_id 最大者。

    返回列：product_code, product_name, val_date, cum_payout, table_unit_nav, table_cum_unit_nav
    其中 table_* 为估值表内同行的单位净值/累计单位净值，仅用于校验，不作为净值来源。
    """
    from fundadmin.clients.schema import get_engine

    engine = engine or get_engine()
    labels = tuple(_LABEL_FIELDS) + _DATE_LABELS
    stmt = text(_SQL_PAYOUT.format(pred=_predicates(labels)))

    records: dict[int, dict] = {}
    with engine.connect() as conn:
        for row in conn.execute(stmt):
            rec = records.get(row.ingest_id)
            if rec is None:
                code = row.db_product_code or _product_from_filename(row.file_name)
                rec = records[row.ingest_id] = {
                    "ingest_id": row.ingest_id,
                    "product_code": code,
                    "file_name": row.file_name,
                    "val_date": None,
                }
            try:
                cells = json.loads(row.cells_json)
            except (TypeError, ValueError):
                continue
            if not isinstance(cells, list) or not cells:
                continue

            label = _normalize_label(cells[0])
            if not label:
                continue
            if label in _DATE_LABELS or _DATE_RE.search(str(cells[0])):
                parsed = _parse_date(cells[0])
                if parsed:
                    rec["val_date"] = parsed

            field = _LABEL_FIELDS.get(label)
            if field:
                value = _to_float(cells[1] if len(cells) > 1 else None)
                if value is not None:
                    rec[field] = value

    rows: list[dict] = []
    for rec in records.values():
        code = rec.get("product_code")
        if code not in PRODCODE_TO_NAME or rec.get("cum_payout") is None:
            continue
        val_date = rec.get("val_date") or _date_from_filename(rec.get("file_name") or "")
        if not val_date:
            continue
        rows.append(
            {
                "product_code": code,
                "product_name": PRODCODE_TO_NAME[code],
                "val_date": val_date,
                "cum_payout": rec["cum_payout"],
                "table_unit_nav": rec.get("unit_nav"),
                "table_cum_unit_nav": rec.get("cum_unit_nav"),
            }
        )

    df = pd.DataFrame(
        rows,
        columns=["product_code", "product_name", "val_date", "cum_payout",
                 "table_unit_nav", "table_cum_unit_nav"],
    )
    if df.empty:
        return df
    df["val_date"] = pd.to_datetime(df["val_date"], errors="coerce")
    df = df.dropna(subset=["val_date"])
    # records 按 ingest_id 升序插入，稳定排序后同 (产品, 日期) 内仍是 ingest_id 升序
    df = df.sort_values(["product_code", "val_date"], kind="stable")
    df = df.drop_duplicates(subset=["product_code", "val_date"], keep="last")
    return df.reset_index(drop=True)


def load_nav_history(*, engine: Engine | None = None) -> pd.DataFrame:
    """产品净值序列（口径已校验，来自 product_valuation）。"""
    from fundadmin.clients.schema import get_engine

    engine = engine or get_engine()
    with engine.connect() as conn:
        df = pd.read_sql(text(_SQL_NAV), conn, parse_dates=["as_of_date"])
    return df.rename(columns={"as_of_date": "val_date"})


def _payout_forward_filled(nav: pd.DataFrame, payout: pd.DataFrame) -> pd.DataFrame:
    """把累计派现贴到净值序列上；缺口按产品前向填充。

    累计派现是单调不减量，估值表并非每个交易日都带该字段，
    因此缺口必须 ffill；若直接按 0 处理会把"字段缺失"误判成"当日分红"。
    """
    if payout is None or payout.empty:
        merged = nav.copy()
        merged["cum_payout"] = 0.0
        return merged

    merged = nav.merge(
        payout[["product_code", "val_date", "cum_payout"]],
        on=["product_code", "val_date"],
        how="left",
    )
    merged["cum_payout"] = merged.groupby("product_code")["cum_payout"].ffill().fillna(0.0)
    return merged


def detect_dividend_events(
    series: pd.DataFrame, *, min_amount: float = 1e-6, nav_window: pd.DataFrame | None = None
) -> pd.DataFrame:
    """由累计派现的增量还原单次分红事件（除息日）。

    事件在**派现序列本身**上求差分，而不是在贴回净值后的序列上，
    否则窗口起点会把"窗口之前就已派完的累计分红"误报成当日事件。

    参数:
        series: 含 product_code / val_date / cum_payout 的派现序列。
        min_amount: 判定为分红的最小每份金额，用于过滤浮点噪声。
        nav_window: 含 product_code / val_date / unit_nav 的净值序列；提供时会标注
            事件是否落在净值窗口内，以及当日净值跌幅能否被该笔分红解释。

    返回列：product_code, product_name, val_date, amount_per_unit, cum_payout_after,
            in_nav_window[, unit_nav_before, unit_nav, nav_change, explained_by_payout]

    注:
        序列首个观测点的前值未知，此处按 0 处理；若某产品的派现序列从非零值开始，
        首行会被记为一笔分红。是否为真需结合 `explained_by_payout` 判断
        （分红当日净值应同步下跌约 amount/nav_before）。
    """
    if series.empty:
        return pd.DataFrame(
            columns=["product_code", "product_name", "val_date",
                     "amount_per_unit", "cum_payout_after", "in_nav_window"]
        )

    work = series.sort_values(["product_code", "val_date"]).copy()
    prev = work.groupby("product_code")["cum_payout"].shift(1).fillna(0.0)
    work["amount_per_unit"] = work["cum_payout"] - prev

    hits = work[work["amount_per_unit"] > min_amount].copy()
    hits = hits.rename(columns={"cum_payout": "cum_payout_after"})
    hits["in_nav_window"] = True
    hits["nav_change"] = None
    hits["explained_by_payout"] = None

    if nav_window is not None and not nav_window.empty:
        nav_only = nav_window[["product_code", "val_date", "unit_nav"]]
        prev_nav = nav_only.sort_values(["product_code", "val_date"]).copy()
        prev_nav["unit_nav_before"] = prev_nav.groupby("product_code")["unit_nav"].shift(1)
        merged = hits.merge(prev_nav, on=["product_code", "val_date"], how="left")

        merged["nav_change"] = merged["unit_nav"] / merged["unit_nav_before"] - 1.0
        # 分红当日净值应下跌约 amount/nav_before；两者吻合即视为被分红解释
        expected = -merged["amount_per_unit"] / merged["unit_nav_before"]
        merged["explained_by_payout"] = (merged["nav_change"] - expected).abs() <= 0.01
        merged["in_nav_window"] = merged["unit_nav_before"].notna()
        merged["nav_change"] = merged["nav_change"].where(merged["in_nav_window"])
        merged["explained_by_payout"] = merged["explained_by_payout"].where(
            merged["in_nav_window"]
        )
        return merged[
            ["product_code", "product_name", "val_date", "amount_per_unit",
             "cum_payout_after", "in_nav_window", "unit_nav_before", "unit_nav",
             "nav_change", "explained_by_payout"]
        ].reset_index(drop=True)

    return hits[
        ["product_code", "product_name", "val_date", "amount_per_unit",
         "cum_payout_after", "in_nav_window"]
    ].reset_index(drop=True)


def build_reinvested_nav(series: pd.DataFrame) -> pd.DataFrame:
    """分红再投资复权序列（按份额法，与算术加回口径对照）。

    算术加回（`build_adjusted_nav` 的 adj_nav = 单位净值 + 累计派现）把每笔分红
    按**面值**加回，没有考虑这笔钱本身在除息日之后还能继续赚钱；
    分红再投资则是拿这笔钱在除息净值上买入新份额，份额此后随净值一起增长。

    逐日递推（持有份额 shares，初值 1，单位净值即 NAV）：
        当日分红        d_t = 累计派现(t) - 累计派现(t-1)
        除权日当日回报  r_t = (NAV(t) + d_t) / NAV(t-1) - 1        # 已剔除分红本身
        分红再投资      shares *= (1 + d_t / NAV(t))                # 分红在除息净值上买入新份额
        组合市值        value(t) = shares * NAV(t)

    注意: shares 只在分红时变化。市场涨跌已由 NAV(t) 体现，
    不可再让 shares 复利增长，否则 value = shares * NAV 会重复计入净值涨幅。

    收益指数以首日市值为 1.0 归一。
    与算术加回的差异只来自"分红资金在除息后是否继续复利"，两种口径不可混用。

    返回列：在 build_adjusted_nav 输出基础上追加
        dividend_on_day, day_total_return, reinvest_shares,
        reinvest_nav, reinvest_return_index
    """
    if series.empty:
        return series.copy()

    work = series.sort_values(["product_code", "val_date"]).copy()
    work["cum_payout"] = work["cum_payout"].fillna(0.0)

    rows: list[dict] = []
    for _, group in work.groupby("product_code", sort=False):
        group = group.sort_values("val_date")
        base_nav = float(group["unit_nav"].iloc[0])
        # 首行没有前值：累计派现按"已含"处理，当日分红记 0，
        # 否则窗口之前就派完的分红会在首日被当成当日分红再投资一次。
        prev_cum = float(group["cum_payout"].iloc[0])
        prev_nav = None
        shares = 1.0
        for row in group.itertuples():
            nav = float(row.unit_nav)
            cum = float(row.cum_payout)
            # 累计派现单调不减；负向跳变视为数据噪声，按 0 处理
            dividend = max(cum - prev_cum, 0.0) if prev_nav is not None else 0.0
            if prev_nav is None:
                day_return = None
            else:
                day_return = (nav + dividend) / prev_nav - 1.0
            shares *= 1.0 + dividend / nav
            rows.append(
                {
                    "product_code": row.product_code,
                    "product_name": row.product_name,
                    "val_date": row.val_date,
                    "unit_nav": nav,
                    "cum_payout": cum,
                    "dividend_on_day": dividend,
                    "day_total_return": day_return,
                    "reinvest_shares": shares,
                    "reinvest_nav": shares * nav,
                    "reinvest_return_index": (shares * nav) / base_nav,
                }
            )
            prev_cum, prev_nav = cum, nav

    out = pd.DataFrame(rows)
    return out[
        ["product_code", "product_name", "val_date", "unit_nav", "cum_payout",
         "dividend_on_day", "day_total_return", "reinvest_shares",
         "reinvest_nav", "reinvest_return_index"]
    ].reset_index(drop=True)


def build_nav_variants(
    nav: pd.DataFrame | None = None,
    payout: pd.DataFrame | None = None,
    *,
    engine: Engine | None = None,
) -> pd.DataFrame:
    """把三种口径合并到一张表，便于对照与出图。

    返回列：product_code, product_name, val_date, unit_nav, cum_payout,
            dividend_on_day, adj_nav（算术加回）, total_return_index（算术加回）,
            day_total_return, reinvest_shares, reinvest_nav, reinvest_return_index

    三种口径的区间收益定义不同，不可混用：
    - 未复权: 纯单位净值，不含任何分红
    - 算术加回: (NAV+累计派现) 的比值；分红按面值加回，且基期若含历史派现会被稀释
    - 分红再投资: 份额法，分红在除息净值上买入新份额并继续复利
    """
    if nav is None:
        nav = load_nav_history(engine=engine)
    if payout is None:
        payout = load_payout_series(engine=engine)

    merged = _payout_forward_filled(nav, payout)
    keys = ["product_code", "product_name", "val_date", "unit_nav", "cum_payout"]
    arith = build_adjusted_nav(nav, payout)[keys + ["adj_nav", "total_return_index", "daily_total_return"]]
    rein = build_reinvested_nav(merged[keys])
    return arith.merge(
        rein.drop(columns=["product_name", "cum_payout"]),
        on=["product_code", "val_date", "unit_nav"],
        how="left",
    ).sort_values(["product_code", "val_date"]).reset_index(drop=True)


def build_adjusted_nav(
    nav: pd.DataFrame, payout: pd.DataFrame | None = None, *, engine: Engine | None = None
) -> pd.DataFrame:
    """生成算术加回复权序列（累计净值口径）。

    复权净值 = 单位净值 + 累计派现（与券商"累计单位净值"恒等）。
    `total_return_index` 为以首日为基点的累计收益指数。

    注意: 这是**分红按面值加回**的口径，不含分红资金的再投资复利。
    需要与投资者实际收益一致的分红再投资口径，请用 build_reinvested_nav。

    参数:
        nav: 含 product_code / val_date / unit_nav 的净值序列；为 None 时从库中读取。
        payout: 含 product_code / val_date / cum_payout；为 None 时从库中读取。

    返回列：product_code, product_name, val_date, unit_nav, cum_payout,
            adj_nav, total_return_index, daily_total_return
    """
    if nav is None:
        nav = load_nav_history(engine=engine)
    if payout is None:
        payout = load_payout_series(engine=engine)

    work = _payout_forward_filled(nav, payout)
    work = work.sort_values(["product_code", "val_date"]).copy()
    work["adj_nav"] = work["unit_nav"] + work["cum_payout"]

    grp = work.groupby("product_code", sort=False)
    # 分红当日回报必须用**当日分红**（累计派现之差），而不是累计派现本身；
    # 后者会把历史已派金额重复计入，导致除息日收益被高估。
    day_dividend = grp["cum_payout"].diff().fillna(0.0).clip(lower=0.0)
    prev_nav = grp["unit_nav"].shift(1)
    work["daily_total_return"] = (work["unit_nav"] + day_dividend) / prev_nav - 1.0
    work["daily_total_return"] = work["daily_total_return"].where(
        prev_nav.notna() & (prev_nav != 0)
    )

    base = grp["adj_nav"].transform("first")
    work["total_return_index"] = work["adj_nav"] / base
    return work[
        ["product_code", "product_name", "val_date", "unit_nav", "cum_payout",
         "adj_nav", "total_return_index", "daily_total_return"]
    ].reset_index(drop=True)
