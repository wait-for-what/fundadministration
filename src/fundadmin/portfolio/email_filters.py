"""基金持仓邮件范围过滤。

用途:
- 为 fund_portfolio 的 IMAP 同步限定产品相关邮件和附件，减少无关研究邮件、
  策略名单和其他 Excel 附件下载。
- 同时为 IMAP 服务端 SEARCH 构造 SUBJECT / FROM / DATE 联合表达式，把
  过滤尽量推到服务端，从源头减少不必要的邮件传输与落盘。

输入:
- 邮件主题、发件人、附件文件名（客户端判定）。
- 抓取日期窗口、可选发件人白名单、可选主题关键词（服务端 SEARCH 构造）。

输出:
- 布尔判断：是否属于当前产品报表链路需要的邮件或附件。
- IMAP SEARCH 字符串：可直接传给 `imaplib.IMAP4*.search`。

失败行为:
- 输入为空或无法识别时返回 False；调用方可通过 `--no-product-scope` 回退到全量行为。
- 服务端 SEARCH 若被某些老服务拒绝（罕见），调用方应回退到 DATE-only 查询并
  打 WARN，本模块只负责构造合法表达式不负责回退。

调用示例:
- `is_target_fund_email(subject="【中信证券】2026-05-07 104902--HKD--EQ...")`
- `is_target_fund_attachment(filename="SCD704_弘运盛泰全球视野...资产估值表.xls")`
- `build_imap_search_criteria(since=date(2026,5,11), before=date(2026,5,12))`
"""

from __future__ import annotations

from collections.abc import Iterable
from datetime import date
from functools import lru_cache

from fundadmin.portfolio.cross_broker_report import PRODUCT_CONFIG

_REPORT_TOKENS = {
    "statement",
    "估值",
    "估值日",
    "估值表",
    "估值报告",
    "履约保障",
    "收益互换",
    "资产估值表",
    "指标计算",
    "持仓",
}
_EXCLUDE_IF_NO_PRODUCT = {
    "cicc strategy",
    "sector top picks",
    "covered list",
    "中金策略",
    "覆盖标的",
    "覆盖公司",
}


def _normalize(value: str) -> str:
    return str(value or "").strip().lower().replace(" ", "")


@lru_cache(maxsize=1)
def _product_tokens() -> tuple[str, ...]:
    tokens: set[str] = {"弘运盛泰"}
    for cfg in PRODUCT_CONFIG:
        name = str(cfg.get("name", "") or "").strip()
        if name:
            tokens.add(name)
            tokens.add(f"弘运盛泰{name}")
        for key in ("citic_codes", "cicc_codes"):
            for code in cfg.get(key, []) or []:
                text = str(code or "").strip()
                if text:
                    tokens.add(text)
    return tuple(sorted({_normalize(token) for token in tokens if token}, key=len, reverse=True))


def _has_product_token(text: str) -> bool:
    normalized = _normalize(text)
    return any(token in normalized for token in _product_tokens())


def _has_report_token(text: str) -> bool:
    normalized = _normalize(text)
    return any(_normalize(token) in normalized for token in _REPORT_TOKENS)


def is_target_fund_email(*, subject: str, sender: str = "") -> bool:
    """判断邮件头是否属于基金持仓/估值同步范围。"""
    text = f"{subject} {sender}"
    normalized = _normalize(text)
    has_product = _has_product_token(text)
    if has_product:
        return True
    if any(_normalize(token) in normalized for token in _EXCLUDE_IF_NO_PRODUCT):
        return False
    return "弘运盛泰" in normalized and _has_report_token(text)


def is_target_fund_attachment(*, filename: str, subject: str = "") -> bool:
    """判断附件是否属于基金持仓/估值同步范围。"""
    return is_target_fund_email(subject=f"{subject} {filename}")


# ---- 以下为 IMAP 服务端 SEARCH 表达式构造 ---------------------------------


def _ascii_search_tokens() -> tuple[str, ...]:
    """从 _product_tokens() 中筛出可在 IMAP 服务端安全搜索的 ASCII token。

    约束:
    - IMAP SEARCH SUBJECT 在多数服务端只对 7-bit ASCII 做可靠匹配，CJK 需要
      UTF-8 CHARSET 才能搜，兼容性差。这里只下推 ASCII token（如
      `104902`、`scd704`），中文 token（`全球视野`、`弘运盛泰`）走客户端
      `is_target_fund_email` 兜底，保证落盘正确性。
    - 至少包含一个字母数字字符，避免误把空白或符号 token 推下去。

    返回:
    - 已去重、按长度倒序排列的 ASCII token 元组。
    """
    tokens: list[str] = []
    for token in _product_tokens():
        if not token:
            continue
        try:
            token.encode("ascii")
        except UnicodeEncodeError:
            continue
        if not any(ch.isalnum() for ch in token):
            continue
        tokens.append(token)
    return tuple(sorted(set(tokens), key=len, reverse=True))


def _imap_search_quote(value: str) -> str:
    """对 IMAP SEARCH 字符串做最小转义。

    IMAP astring 允许双引号内出现普通字符；`\\` 与 `"` 需要反斜杠转义。
    这里保守做法：剥掉控制字符，转义反斜杠和双引号。
    """
    cleaned = "".join(ch for ch in str(value or "") if 32 <= ord(ch) or ord(ch) >= 128)
    return cleaned.replace("\\", "\\\\").replace('"', '\\"')


def _or_chain(field: str, tokens: Iterable[str]) -> str:
    """构造 IMAP `OR FIELD "a" FIELD "b" ...` 链。

    边界:
    - 空 tokens 返回空串。
    - 1 个 token 直接返回 `FIELD "x"`，无 OR。
    - 多个 token 右关联嵌套：`OR FIELD a OR FIELD b FIELD c`，符合 RFC 3501
      `OR` 二元定义。
    """
    cleaned: list[str] = []
    seen: set[str] = set()
    for token in tokens:
        text = str(token or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        cleaned.append(text)
    if not cleaned:
        return ""
    if len(cleaned) == 1:
        return f'{field} "{_imap_search_quote(cleaned[0])}"'
    expr = f'{field} "{_imap_search_quote(cleaned[-1])}"'
    for token in reversed(cleaned[:-1]):
        expr = f'OR {field} "{_imap_search_quote(token)}" {expr}'
    return expr


_IMAP_MONTHS = (
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
)


def _imap_date_str(value: date) -> str:
    """IMAP SEARCH 日期格式: dd-Mon-yyyy（RFC 3501）。

    使用硬编码英文月份缩写而非 `strftime("%b")`，避免依赖系统 locale。
    """
    return f"{value.day:02d}-{_IMAP_MONTHS[value.month - 1]}-{value.year:04d}"


def build_imap_search_criteria(
    *,
    since: date,
    before: date,
    sender_allowlist: Iterable[str] = (),
    subject_keywords: Iterable[str] = (),
    scope_from_products: bool = True,
) -> str:
    """构造 IMAP SEARCH 表达式，把过滤尽量推到服务端，减少客户端流量。

    返回形如::

        (SINCE "11-May-2026" BEFORE "12-May-2026" OR SUBJECT "scd704" SUBJECT "104902" FROM "broker@citic.com")

    用途:
    - `since` / `before` 必填，对应 IMAP RFC 3501 的 SEARCH SINCE / BEFORE。
    - `scope_from_products=True` 时把 `_ascii_search_tokens()` 拼入 SUBJECT
      OR 链；中文 token 留给客户端 `is_target_fund_email` 兜底。
    - `subject_keywords` 中的 ASCII 关键词会合并进同一个 SUBJECT OR 链；
      非 ASCII 项会被静默丢弃（客户端仍会按子串过滤）。
    - `sender_allowlist` 拼成 OR FROM 链。
    - 多组之间用 IMAP 隐式 AND（空格）连接，符合现有 cli.py 拼装风格。

    返回:
    - 形如 `(...)` 的 IMAP SEARCH 字符串，可直接传给 `imap.search(None, ...)`。
    """
    parts: list[str] = [
        f'SINCE "{_imap_date_str(since)}"',
        f'BEFORE "{_imap_date_str(before)}"',
    ]

    subject_tokens: list[str] = []
    if scope_from_products:
        subject_tokens.extend(_ascii_search_tokens())
    for keyword in subject_keywords or ():
        text = str(keyword or "").strip()
        if not text:
            continue
        try:
            text.encode("ascii")
        except UnicodeEncodeError:
            # 中文等非 ASCII 关键词不推服务端，客户端 keywords 仍会作子串过滤
            continue
        subject_tokens.append(text.lower())
    subject_expr = _or_chain("SUBJECT", subject_tokens)
    if subject_expr:
        parts.append(subject_expr)

    from_expr = _or_chain("FROM", sender_allowlist or ())
    if from_expr:
        parts.append(from_expr)

    return f"({' '.join(parts)})"
