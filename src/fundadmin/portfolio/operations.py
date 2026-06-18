"""基金持仓 Excel 拉取与组合汇总 CLI（god module - 治理 2026-04 P2.7）。

治理提示:
    本文件 1710 行，CLI 入口不薄；已登记在
    tests.unit.test_project_standards.KNOWN_GOD_MODULES。建议拆分方向：
    - cli.py: 仅命令注册（保留 typer/click 入口）。
    - operations.py: 业务函数（拉 Excel、解析、汇总、推送）。
    - reports.py: 报表/邮件/HTML 渲染。
    拆分后 CLI 改为薄壳调 operations，符合 .claude/rules/paths/20_cli_jobs_pipelines.md
    的"CLI 是入口、业务下沉到 src"约束。

用途:
- 通过 IMAP 拉取指定交易日的基金邮件附件。
- 解析持仓表与估值表 Excel，并按公司维度汇总组合暴露。

输入:
- IMAP 账号参数，优先从命令行读取，也支持 `.env` 中的 `IMAP_*` 配置。
- `email-sync` 子命令的交易日、发件人白名单、主题关键词。
- `build` 子命令的附件目录，或默认读取 `outputs/excels/fund_inbox/<trade_date>/`。

输出:
- 下载的 Excel 附件写入 `outputs/excels/fund_inbox/<trade_date>/`。
- 汇总报表写入 `outputs/reports/fund_portfolios/portfolio_<trade_date>.xlsx`。

失败行为:
- 日期格式非法、IMAP 登录失败、附件目录不存在、关键表头缺失时抛出异常并返回非零退出码。

调用示例:
- `python -m apps.cli ops fund-portfolio email-sync --trade-date 2026-04-17`
- `python -m apps.cli ops fund-portfolio build --trade-date 2026-04-17`
"""

from __future__ import annotations

import argparse
import hashlib
import imaplib
import json
import logging
import re
import shutil
import ssl
import time
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from email import message_from_bytes
from email.header import decode_header
from email.message import Message
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

import pandas as pd

from fundadmin.core.config import get_env, load_env
from fundadmin.core.paths import repo_root
from fundadmin.portfolio.client_notifier import send_client_nav_emails
from fundadmin.portfolio.cross_broker_report import (
    PRODUCT_CONFIG,
    CrossBrokerInput,
    build_cross_broker_report,
    build_product_reports,
    build_summary_excel,
    score_product_inputs_for_date,
)
from fundadmin.portfolio.email_filters import (
    build_imap_search_criteria,
    is_target_fund_attachment,
    is_target_fund_email,
)
from fundadmin.portfolio.maintenance import (
    format_bytes,
    prune_inbox,
)
from fundadmin.portfolio.by_broker_email import send_by_broker_summary_email
from fundadmin.portfolio.notifier import send_matrix_email
from fundadmin.portfolio.viz import generate_portfolio_pie_chart
from fundadmin.notifications.email import SmtpConfig

logger = logging.getLogger(__name__)

DEFAULT_SYNC_LOOKBACK = 1
# 发布段回看天数：券商报告 T+1/T+2 才到邮箱，需扫描近若干天收件目录，
# 按文件名里的交易日重新分组，确保较旧但补齐的交易日也能被发出。
PUBLISH_SCAN_DAYS = 10
DEFAULT_STATE_FILENAME = "email_sync_state.json"
PUBLISH_STATE_FILENAME = "published_state.json"
SUPPORTED_ATTACHMENT_SUFFIXES = (".xlsx", ".xlsm", ".xls", ".csv")
FILENAME_ISO_DATE_PATTERN = re.compile(r"(?<!\d)(20\d{2}-\d{2}-\d{2})(?!\d)")
FILENAME_COMPACT_DATE_PATTERN = re.compile(r"(?<!\d)(20\d{6})(?!\d)")

HOLDINGS_NAME_LABELS = {
    "证券名称",
    "股票名称",
    "证券简称",
    "股票简称",
}
HOLDINGS_CODE_LABELS = {
    "证券代码",
    "股票代码",
    "代码",
}
HOLDINGS_MARKET_VALUE_LABELS = {
    "市值",
    "市值(元)",
    "市值（元）",
    "持仓市值",
    "持仓市值(元)",
    "持仓市值（元）",
}
HOLDINGS_WEIGHT_LABELS = {
    "占净值比例",
    "占基金净值比例",
    "占净资产比例",
    "市值占比",
}
VALUATION_NAV_LABELS = {
    "单位净值",
    "基金单位净值",
    "份额净值",
    "净值",
}
VALUATION_NET_ASSETS_LABELS = {
    "基金资产净值",
    "资产净值",
    "基金净资产",
    "期末基金资产净值",
    "期末资产净值",
}
VALUATION_FILENAME_KEYWORDS = ("估值", "valuation")
HOLDINGS_FILENAME_KEYWORDS = ("持仓", "position", "holding")


@dataclass(frozen=True)
class ImapConfig:
    host: str
    user: str
    password: str
    mailbox: str = "INBOX"
    port: int = 993
    use_ssl: bool = True


@dataclass(frozen=True)
class FundValuation:
    nav: float | None
    net_assets: float | None


@dataclass(frozen=True)
class EmailSyncResult:
    trade_date: date
    out_dir: Path
    saved_paths: tuple[Path, ...]
    matched_messages: int
    processed_messages: int
    skipped_messages: int
    state_file: Path | None = None


def _parse_ymd(value: str) -> date:
    text = str(value or "").strip()
    if not text:
        raise ValueError("date is empty")
    return datetime.strptime(text, "%Y-%m-%d").date()


def _parse_optional_ymd(value: str) -> date | None:
    text = str(value or "").strip()
    if not text:
        return None
    return _parse_ymd(text)


def _default_inbox_root() -> Path:
    return repo_root() / "outputs" / "excels" / "fund_inbox"


def _default_inbox_dir(trade_date: date) -> Path:
    return _default_inbox_root() / trade_date.isoformat()


def _default_report_root() -> Path:
    return repo_root() / "outputs" / "reports" / "fund_portfolios"


def _default_portfolio_report_path(trade_date: date) -> Path:
    return _default_report_root() / f"portfolio_{trade_date.isoformat()}.xlsx"


def _default_mail_state_path(inbox_root: Path | None = None) -> Path:
    root = Path(inbox_root) if inbox_root is not None else _default_inbox_root()
    return root / "_state" / DEFAULT_STATE_FILENAME


def _default_publish_state_path(inbox_root: Path | None = None) -> Path:
    root = Path(inbox_root) if inbox_root is not None else _default_inbox_root()
    return root / "_state" / PUBLISH_STATE_FILENAME


def _extract_trade_dates_from_text(text: str) -> list[date]:
    resolved: set[date] = set()
    raw = str(text or "")
    for match in FILENAME_ISO_DATE_PATTERN.findall(raw):
        try:
            resolved.add(_parse_ymd(match))
        except ValueError:
            pass
    for match in FILENAME_COMPACT_DATE_PATTERN.findall(raw):
        try:
            resolved.add(datetime.strptime(match, "%Y%m%d").date())
        except ValueError:
            pass
    return sorted(resolved)


def _extract_trade_dates_from_path(path: Path) -> list[date]:
    return _extract_trade_dates_from_text(path.name)


def _fallback_business_trade_dates(*, asof_date: date, lookback: int) -> list[date]:
    target = max(1, int(lookback))
    current = asof_date
    resolved: list[date] = []
    while len(resolved) < target:
        if current.weekday() < 5:
            resolved.append(current)
        current -= timedelta(days=1)
    return list(sorted(resolved))


def _resolve_latest_trade_dates(
    *, asof_date: date | None = None, lookback: int = 1
) -> tuple[list[date], str | None]:
    """解析最近 N 个交易日。

    约束:
    - 本仓库自包含、不连交易日历库，按工作日（周一至周五）回退，不剔除节假日。
    - 需要精确交易日时，命令行显式传入 --date。
    """
    target = asof_date or date.today()
    return _fallback_business_trade_dates(asof_date=target, lookback=max(1, int(lookback))), None


def _resolve_trade_date_arg(value: str, *, today: date | None = None) -> date:
    text = str(value or "").strip()
    if not text:
        raise ValueError("date is empty")
    if text.lower() != "latest":
        return _parse_ymd(text)
    resolved, warning = _resolve_latest_trade_dates(asof_date=today or date.today(), lookback=1)
    if warning:
        print(f"[WARN] {warning}")
    if not resolved:
        raise RuntimeError("cannot resolve latest trade date")
    return resolved[-1]


def _ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _decode_email_header(value: str | None) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for chunk, encoding in decode_header(value):
        if isinstance(chunk, bytes):
            try:
                parts.append(chunk.decode(encoding or "utf-8", errors="replace"))
            except Exception:
                parts.append(chunk.decode("utf-8", errors="replace"))
        else:
            parts.append(str(chunk))
    return "".join(parts).strip()


def _safe_filename(name: str) -> str:
    text = re.sub(r"[\\/:*?\"<>|]+", "_", str(name or "").strip())
    text = re.sub(r"\s+", " ", text).strip()
    return text[:160] if len(text) > 160 else text


def _parse_csv_arg(value: str) -> list[str]:
    return [item.strip() for item in str(value or "").split(",") if item.strip()]


def _load_email_sync_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"version": 1, "processed_messages": {}}
    raw = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError(f"invalid state file payload: {path}")
    processed = raw.get("processed_messages")
    if not isinstance(processed, dict):
        raw["processed_messages"] = {}
    raw["version"] = int(raw.get("version") or 1)
    return raw


def _save_email_sync_state(path: Path, state: dict[str, Any]) -> None:
    _ensure_dir(path.parent)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_publish_state(path: Path) -> dict[str, Any]:
    """发布记录：trade_date -> {"matrix_sent_at": iso, "clients_sent_at": iso}。

    用于 sync-latest 发布段去重：同一交易日某渠道已成功发送后不再重发，
    数据不齐的交易日不会写入记录，留待后续补料后自动补发。
    """
    if not path.exists():
        return {"version": 1, "published": {}}
    raw = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError(f"invalid publish state payload: {path}")
    if not isinstance(raw.get("published"), dict):
        raw["published"] = {}
    raw["version"] = int(raw.get("version") or 1)
    return raw


def _save_publish_state(path: Path, state: dict[str, Any]) -> None:
    _ensure_dir(path.parent)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _publish_channel_done(state: dict[str, Any], trade_date: date, channel: str) -> bool:
    entry = state.get("published", {}).get(trade_date.isoformat())
    if not isinstance(entry, dict):
        return False
    return bool(entry.get(f"{channel}_sent_at"))


def _mark_publish_channel(state: dict[str, Any], trade_date: date, channel: str) -> None:
    published = state.setdefault("published", {})
    entry = published.setdefault(trade_date.isoformat(), {})
    entry[f"{channel}_sent_at"] = datetime.now().isoformat(timespec="seconds")


def _message_key_from_header(msg: Message, header_bytes: bytes) -> str:
    message_id = str(msg.get("Message-ID") or msg.get("Message-Id") or "").strip()
    if message_id:
        return f"message-id:{message_id.lower()}"
    return f"header-sha1:{hashlib.sha1(bytes(header_bytes)).hexdigest()}"


def _message_token(message_key: str) -> str:
    return hashlib.sha1(str(message_key).encode("utf-8")).hexdigest()[:12]


def _build_state_entry(
    *,
    message_key: str,
    msg: Message,
    trade_date: date,
    saved_paths: list[Path],
) -> dict[str, Any]:
    return {
        "message_key": message_key,
        "message_id": str(msg.get("Message-ID") or msg.get("Message-Id") or "").strip(),
        "subject": _decode_email_header(msg.get("Subject")),
        "sender": _decode_email_header(msg.get("From")),
        "trade_date": trade_date.isoformat(),
        "saved_files": [path.name for path in saved_paths],
        "processed_at_utc": datetime.utcnow().replace(microsecond=0).isoformat() + "Z",
    }


def _imap_date(value: date) -> str:
    return value.strftime("%d-%b-%Y")


_IMAP_INTERNALDATE_MON = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}


def _parse_imap_internaldate(meta: bytes | str | None) -> date | None:
    """从 FETCH 元信息里解析 INTERNALDATE 的日期部分（收件日）。

    QQ/exmail 的 IMAP SEARCH 会忽略 SINCE/BEFORE 而返回整箱，因此需要在客户端
    用 INTERNALDATE 自行按收件日筛选，避免把全箱历史都当成"当日"邮件抓下来。
    """
    if meta is None:
        return None
    text = meta.decode("ascii", "replace") if isinstance(meta, (bytes, bytearray)) else str(meta)
    m = re.search(r'INTERNALDATE "(\d{1,2})-(\w{3})-(\d{4})', text)
    if not m:
        return None
    try:
        return date(int(m.group(3)), _IMAP_INTERNALDATE_MON.get(m.group(2), 1), int(m.group(1)))
    except ValueError:
        return None


def _imap_search_with_fallback(
    client: imaplib.IMAP4,
    criteria: str,
    *,
    since: date,
    before: date,
    sender_tokens: list[str],
) -> list[bytes]:
    """执行 IMAP SEARCH，若服务端拒绝 SUBJECT 范围则回退到 DATE+FROM 查询。

    边界:
    - 部分老 IMAP 服务对深嵌套 OR 或非常规 SUBJECT 子句返回 NO/BAD。
    - 这里捕获两种失败信号：search 返回 status != OK，或 imaplib 抛
      `imaplib.IMAP4.error`。
    - 回退查询保证 DATE 与 FROM 仍然下推，最差也只是返回当日全部邮件 ID。
    """
    try:
        status, data = client.search(None, criteria)
    except imaplib.IMAP4.error as exc:  # type: ignore[attr-defined]
        status = "BAD"
        data = []
        print(f"[WARN] IMAP SEARCH rejected ({exc}); falling back to DATE+FROM only")
    if status != "OK":
        print(
            "[WARN] server rejected SUBJECT scope; "
            "falling back to client-side filter (DATE+FROM only)"
        )
        fallback = build_imap_search_criteria(
            since=since,
            before=before,
            sender_allowlist=sender_tokens,
            subject_keywords=(),
            scope_from_products=False,
        )
        status, data = client.search(None, fallback)
        if status != "OK":
            return []
    if not data or not data[0]:
        return []
    return [item for item in data[0].split() if item]


def _iter_attachments(msg: Message) -> Iterable[tuple[str, bytes]]:
    for part in msg.walk():
        if part.is_multipart():
            continue
        disposition = str(part.get("Content-Disposition") or "").lower()
        if "attachment" not in disposition:
            continue
        filename = _decode_email_header(part.get_filename())
        payload = part.get_payload(decode=True)
        if not filename or payload is None:
            continue
        if not isinstance(payload, (bytes, bytearray)):
            continue
        yield filename, bytes(payload)


def _normalize_label(value: Any) -> str:
    text = str(value or "").strip()
    if not text or text.lower() == "nan":
        return ""
    text = text.replace("\n", "").replace("\r", "").replace("\t", "")
    return text


def fetch_excel_attachments_via_imap(
    *,
    imap: ImapConfig,
    target_date: date,
    out_dir: Path,
    sender_allowlist: list[str] | None = None,
    subject_keywords: list[str] | None = None,
    product_scope: bool = True,
    scope_from_products: bool = True,
) -> list[Path]:
    _ensure_dir(out_dir)
    allow_senders = {str(x or "").strip().lower() for x in (sender_allowlist or []) if str(x or "").strip()}
    keywords = [str(x or "").strip().lower() for x in (subject_keywords or []) if str(x or "").strip()]

    context = ssl.create_default_context()
    client: imaplib.IMAP4
    if imap.use_ssl:
        client = imaplib.IMAP4_SSL(imap.host, int(imap.port), ssl_context=context)
    else:
        client = imaplib.IMAP4(imap.host, int(imap.port))
    try:
        client.login(imap.user, imap.password)
        select_status, _ = client.select(imap.mailbox)
        if select_status != "OK":
            # 部分中文邮箱的收件箱名称为中文，尝试 fallback
            fallback = "收件箱" if imap.mailbox.upper() == "INBOX" else "INBOX"
            select_status, _ = client.select(fallback)
            if select_status != "OK":
                raise RuntimeError(f"无法选择邮箱文件夹: {imap.mailbox} / {fallback}")

        since = target_date
        before = target_date + timedelta(days=1)
        # 构造 SEARCH 表达式：DATE + 可选 SUBJECT 产品 token + 可选 FROM 白名单
        criteria = build_imap_search_criteria(
            since=since,
            before=before,
            sender_allowlist=sorted(allow_senders),
            subject_keywords=keywords,
            scope_from_products=scope_from_products,
        )
        ids = _imap_search_with_fallback(client, criteria, since=since, before=before, sender_tokens=sorted(allow_senders))
        if not ids:
            return []

        def _batch_fetch(client: imaplib.IMAP4, msg_ids: list[bytes], spec: str) -> dict[bytes, bytes]:
            """批量 fetch，返回 {msg_id: payload_bytes}。"""
            if not msg_ids:
                return {}
            # imaplib 要求 id 序列是逗号分隔的 ASCII 字符串
            seq = b",".join(msg_ids).decode("ascii")
            status, data = client.fetch(seq, spec)
            if status != "OK" or not data:
                return {}
            result: dict[bytes, bytes] = {}
            for item in data:
                if not isinstance(item, tuple) or len(item) < 2:
                    continue
                meta, payload = item[0], item[1]
                if not isinstance(meta, (bytes, bytearray)) or not isinstance(payload, (bytes, bytearray)):
                    continue
                # meta 格式: b"123 (BODY[HEADER] {4567}"
                m = re.match(rb"^(\d+)", meta)
                if m:
                    result[m.group(1)] = bytes(payload)
            return result

        def _batch_fetch_rfc822(client: imaplib.IMAP4, msg_ids: list[bytes]) -> dict[bytes, Message]:
            """批量 fetch RFC822，返回 {msg_id: parsed_message}。"""
            payloads = _batch_fetch(client, msg_ids, "(RFC822)")
            return {mid: message_from_bytes(pb) for mid, pb in payloads.items()}

        # 阶段一：批量 peek header，客户端过滤主题/发件人
        matched_ids: list[bytes] = []
        batch_size = 50
        for i in range(0, len(ids), batch_size):
            batch = ids[i : i + batch_size]
            headers = _batch_fetch(client, batch, "(BODY.PEEK[HEADER])")
            for mid, hb in headers.items():
                msg = message_from_bytes(hb)
                subject = _decode_email_header(msg.get("Subject"))
                sender = _decode_email_header(msg.get("From"))
                sender_lower = sender.lower()
                if allow_senders and not any(token in sender_lower for token in allow_senders):
                    continue
                if keywords:
                    subject_lower = subject.lower()
                    if not any(token in subject_lower for token in keywords):
                        continue
                if product_scope and not is_target_fund_email(subject=subject, sender=sender):
                    continue
                matched_ids.append(mid)

        # 阶段二：批量下载匹配邮件的完整内容，提取附件
        saved: list[Path] = []
        for i in range(0, len(matched_ids), batch_size):
            batch = matched_ids[i : i + batch_size]
            full_msgs = _batch_fetch_rfc822(client, batch)
            for mid, msg in full_msgs.items():
                subject = _decode_email_header(msg.get("Subject"))
                for filename, payload in _iter_attachments(msg):
                    if not str(filename).lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
                        continue
                    if product_scope and not is_target_fund_attachment(
                        filename=filename,
                        subject=subject,
                    ):
                        continue
                    prefix = _safe_filename(subject) or "email"
                    base = f"{prefix}_{mid.decode(errors='ignore')}_{_safe_filename(filename)}"
                    path = out_dir / base
                    path.write_bytes(payload)
                    saved.append(path)
        return saved
    finally:
        try:
            client.logout()
        except Exception:
            # IMAP logout 失败仅记录，不影响调用方主流程。
            logger.debug("IMAP client.logout() failed", exc_info=True)


def _fetch_excel_attachments_via_imap_with_state(
    *,
    imap: ImapConfig,
    target_date: date,
    out_dir: Path,
    sender_allowlist: list[str] | None = None,
    subject_keywords: list[str] | None = None,
    state_file: Path | None = None,
    skip_processed: bool = False,
    product_scope: bool = True,
    scope_from_products: bool = True,
    recent_scan_limit: int = 0,
    enforce_received_date: bool = False,
) -> EmailSyncResult:
    """增量拉取目标收件日的附件。

    full-access(全量)邮箱适配:
    - recent_scan_limit>0 时只扫描最新的 N 封(按 seq 取尾部)，避免每次 peek 整箱
      4 万+ 头导致缓慢与 QQ 掉线；最新 N 封覆盖最近若干周，足够日常增量。
    - enforce_received_date=True 时按 INTERNALDATE 客户端筛收件日窗口[since, before)，
      因为 QQ 的 SEARCH 会忽略 SINCE/BEFORE 返回整箱，否则会把全箱历史误当当日抓下。
    """
    _ensure_dir(out_dir)
    state_path = Path(state_file) if state_file is not None else None
    state = _load_email_sync_state(state_path) if state_path is not None else {"version": 1, "processed_messages": {}}
    processed_messages = dict(state.get("processed_messages") or {})
    allow_senders = {str(x or "").strip().lower() for x in (sender_allowlist or []) if str(x or "").strip()}
    keywords = [str(x or "").strip().lower() for x in (subject_keywords or []) if str(x or "").strip()]

    context = ssl.create_default_context()

    def _connect() -> imaplib.IMAP4:
        """建立 IMAP 连接、登录并选中目标邮箱（含中文收件箱 fallback）。"""
        if imap.use_ssl:
            conn: imaplib.IMAP4 = imaplib.IMAP4_SSL(imap.host, int(imap.port), ssl_context=context)
        else:
            conn = imaplib.IMAP4(imap.host, int(imap.port))
        conn.login(imap.user, imap.password)
        sel_status, _ = conn.select(imap.mailbox)
        if sel_status != "OK":
            fallback = "收件箱" if imap.mailbox.upper() == "INBOX" else "INBOX"
            sel_status, _ = conn.select(fallback)
            if sel_status != "OK":
                raise RuntimeError(f"cannot select mailbox: {imap.mailbox} / {fallback}")
        return conn

    # 可变持有：fetch 中途若被服务端断连（exmail/QQ 大批量 FETCH 常见 socket EOF），
    # 重连后替换此引用，保证后续命令落到新连接上。
    client_box: list[imaplib.IMAP4] = [_connect()]
    try:
        since = target_date
        before = target_date + timedelta(days=1)
        criteria = build_imap_search_criteria(
            since=since,
            before=before,
            sender_allowlist=sorted(allow_senders),
            subject_keywords=keywords,
            scope_from_products=scope_from_products,
        )
        ids = _imap_search_with_fallback(
            client_box[0],
            criteria,
            since=since,
            before=before,
            sender_tokens=sorted(allow_senders),
        )
        # full-access 邮箱里 SEARCH 返回整箱，仅扫描最新 N 封(seq 升序，尾部最新)。
        if recent_scan_limit and len(ids) > recent_scan_limit:
            ids = ids[-recent_scan_limit:]
        if not ids:
            return EmailSyncResult(
                trade_date=target_date,
                out_dir=out_dir,
                saved_paths=(),
                matched_messages=0,
                processed_messages=0,
                skipped_messages=0,
                state_file=state_path,
            )

        def _raw_fetch_with_retry(seq: str, spec: str, *, max_retries: int = 3) -> list:
            """单次 FETCH，遇 abort/socket EOF 时退避重连后重试；返回原始 data 列表。"""
            last_exc: Exception | None = None
            for attempt in range(max_retries):
                try:
                    status, data = client_box[0].fetch(seq, spec)
                    if status != "OK":
                        return []
                    return data or []
                except (imaplib.IMAP4.abort, OSError) as exc:
                    last_exc = exc
                    logger.warning(
                        "IMAP fetch aborted (attempt %d/%d): %s; reconnecting",
                        attempt + 1, max_retries, exc,
                    )
                    try:
                        client_box[0].logout()
                    except Exception:
                        logger.debug("IMAP logout on dead connection failed", exc_info=True)
                    time.sleep(min(2 ** attempt, 8))
                    try:
                        client_box[0] = _connect()
                    except Exception:
                        logger.warning("IMAP reconnect failed", exc_info=True)
            raise last_exc if last_exc is not None else imaplib.IMAP4.abort("fetch failed")

        def _batch_fetch(client: imaplib.IMAP4, msg_ids: list[bytes], spec: str) -> dict[bytes, bytes]:
            if not msg_ids:
                return {}
            seq = b",".join(msg_ids).decode("ascii")
            try:
                data = _raw_fetch_with_retry(seq, spec)
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError):
                # 重试仍失败：二分拆批，隔离过大批量或单封问题邮件，避免整次同步崩溃。
                if len(msg_ids) <= 1:
                    logger.warning("IMAP fetch failed for single message, skipping: %r", msg_ids)
                    return {}
                mid = len(msg_ids) // 2
                return {
                    **_batch_fetch(client_box[0], msg_ids[:mid], spec),
                    **_batch_fetch(client_box[0], msg_ids[mid:], spec),
                }
            result: dict[bytes, bytes] = {}
            for item in data:
                if not isinstance(item, tuple) or len(item) < 2:
                    continue
                meta, payload = item[0], item[1]
                if not isinstance(meta, (bytes, bytearray)) or not isinstance(payload, (bytes, bytearray)):
                    continue
                match = re.match(rb"^(\d+)", meta)
                if match:
                    result[match.group(1)] = bytes(payload)
            return result

        def _batch_fetch_rfc822(client: imaplib.IMAP4, msg_ids: list[bytes]) -> dict[bytes, Message]:
            payloads = _batch_fetch(client, msg_ids, "(RFC822)")
            return {mid: message_from_bytes(payload) for mid, payload in payloads.items()}

        def _batch_peek_meta(msg_ids: list[bytes]) -> dict[bytes, tuple[date | None, bytes]]:
            """批量 peek (INTERNALDATE + 头)，返回 mid -> (收件日, 头字节)；带重连/二分降批。"""
            if not msg_ids:
                return {}
            seq = b",".join(msg_ids).decode("ascii")
            try:
                data = _raw_fetch_with_retry(seq, "(INTERNALDATE BODY.PEEK[HEADER])")
            except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError):
                if len(msg_ids) <= 1:
                    return {}
                mid = len(msg_ids) // 2
                return {**_batch_peek_meta(msg_ids[:mid]), **_batch_peek_meta(msg_ids[mid:])}
            out: dict[bytes, tuple[date | None, bytes]] = {}
            for item in data:
                if not isinstance(item, tuple) or len(item) < 2:
                    continue
                meta, payload = item[0], item[1]
                if not isinstance(meta, (bytes, bytearray)) or not isinstance(payload, (bytes, bytearray)):
                    continue
                m = re.match(rb"^(\d+)", meta)
                if m:
                    out[m.group(1)] = (_parse_imap_internaldate(meta), bytes(payload))
            return out

        matched_meta: dict[bytes, dict[str, str]] = {}
        skipped_messages = 0
        # 表头用较大批，正文（含大附件）用较小批，降低单条命令体积以减少服务端断连。
        batch_size = 50
        rfc822_batch_size = 20
        for i in range(0, len(ids), batch_size):
            batch = ids[i : i + batch_size]
            headers = _batch_peek_meta(batch)
            for mid, (received_date, header_bytes) in headers.items():
                # full-access 邮箱: 按收件日客户端筛窗口，避免把全箱历史误当当日邮件抓取。
                if enforce_received_date and not (
                    received_date is not None and since <= received_date < before
                ):
                    continue
                msg = message_from_bytes(header_bytes)
                subject = _decode_email_header(msg.get("Subject"))
                sender = _decode_email_header(msg.get("From"))
                sender_lower = sender.lower()
                if allow_senders and not any(token in sender_lower for token in allow_senders):
                    continue
                if keywords:
                    subject_lower = subject.lower()
                    if not any(token in subject_lower for token in keywords):
                        continue
                if product_scope and not is_target_fund_email(subject=subject, sender=sender):
                    continue
                message_key = _message_key_from_header(msg, header_bytes)
                if skip_processed and message_key in processed_messages:
                    skipped_messages += 1
                    continue
                matched_meta[mid] = {
                    "message_key": message_key,
                    "subject": subject,
                    "sender": sender,
                }

        saved: list[Path] = []
        processed_count = 0
        matched_ids = list(matched_meta.keys())
        for i in range(0, len(matched_ids), rfc822_batch_size):
            batch = matched_ids[i : i + rfc822_batch_size]
            full_msgs = _batch_fetch_rfc822(client_box[0], batch)
            for mid, msg in full_msgs.items():
                meta = matched_meta.get(mid)
                if meta is None:
                    continue
                subject = str(meta.get("subject") or _decode_email_header(msg.get("Subject"))).strip()
                message_key = str(meta.get("message_key") or "").strip() or _message_key_from_header(msg, b"")
                token = _message_token(message_key)
                saved_for_message: list[Path] = []
                for index, (filename, payload) in enumerate(_iter_attachments(msg), start=1):
                    if not str(filename).lower().endswith(SUPPORTED_ATTACHMENT_SUFFIXES):
                        continue
                    if product_scope and not is_target_fund_attachment(
                        filename=filename,
                        subject=subject,
                    ):
                        continue
                    prefix = _safe_filename(subject) or "email"
                    base = f"{prefix}_{token}_{index:02d}_{_safe_filename(filename)}"
                    path = out_dir / base
                    path.write_bytes(payload)
                    saved.append(path)
                    saved_for_message.append(path)
                processed_messages[message_key] = _build_state_entry(
                    message_key=message_key,
                    msg=msg,
                    trade_date=target_date,
                    saved_paths=saved_for_message,
                )
                processed_count += 1

        if state_path is not None:
            state["processed_messages"] = processed_messages
            _save_email_sync_state(state_path, state)
        return EmailSyncResult(
            trade_date=target_date,
            out_dir=out_dir,
            saved_paths=tuple(saved),
            matched_messages=len(matched_ids) + skipped_messages,
            processed_messages=processed_count,
            skipped_messages=skipped_messages,
            state_file=state_path,
        )
    finally:
        try:
            client_box[0].logout()
        except Exception:
            # IMAP logout 失败仅记录，不影响调用方主流程。
            logger.debug("IMAP client.logout() failed", exc_info=True)


def _to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)) and pd.notna(value):
        return float(value)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return None
    text = text.replace(",", "").replace("，", "")

    percent_match = re.match(r"^(-?\d+(?:\.\d+)?)\s*%$", text)
    if percent_match:
        try:
            return float(percent_match.group(1)) / 100.0
        except Exception:
            return None

    try:
        return float(text)
    except Exception:
        return None


def _looks_like_header_row(values: list[Any]) -> bool:
    keys = {_normalize_label(value) for value in values}
    keys.discard("")
    if not keys:
        return False
    target_labels = (
        HOLDINGS_NAME_LABELS
        | HOLDINGS_CODE_LABELS
        | HOLDINGS_MARKET_VALUE_LABELS
        | HOLDINGS_WEIGHT_LABELS
    )
    hits = sum(1 for label in keys if label in target_labels)
    return hits >= 2


def _detect_holdings_header(df_raw: pd.DataFrame) -> int:
    for index in range(min(60, len(df_raw))):
        if _looks_like_header_row(df_raw.iloc[index].tolist()):
            return index
    return 0


def _normalize_holdings_columns(columns: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for column in columns:
        label = _normalize_label(column)
        if label in HOLDINGS_NAME_LABELS:
            mapping[column] = "security_name"
        elif label in HOLDINGS_CODE_LABELS:
            mapping[column] = "security_code"
        elif label in HOLDINGS_MARKET_VALUE_LABELS:
            mapping[column] = "market_value"
        elif label in HOLDINGS_WEIGHT_LABELS:
            mapping[column] = "weight"
    return mapping


def parse_holdings_excel(path: Path, *, sheet: str = "") -> pd.DataFrame:
    df_raw = pd.read_excel(path, sheet_name=(sheet or 0), header=None, dtype=object, engine="openpyxl")
    header_row = _detect_holdings_header(df_raw)
    headers = [_normalize_label(value) for value in df_raw.iloc[header_row].tolist()]
    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")
    df = df.loc[:, [column for column in df.columns if _normalize_label(column)]]

    df = df.rename(columns=_normalize_holdings_columns(list(df.columns)))
    if "security_name" not in df.columns:
        raise ValueError(f"holdings sheet missing security_name: {path}")
    if "market_value" not in df.columns and "weight" not in df.columns:
        raise ValueError(f"holdings sheet missing market_value/weight: {path}")

    out = pd.DataFrame()
    out["company"] = df["security_name"].astype(str).str.strip()
    out["security_code"] = df["security_code"].astype(str).str.strip() if "security_code" in df.columns else ""
    out["market_value"] = df["market_value"].map(_to_number) if "market_value" in df.columns else None
    out["weight"] = df["weight"].map(_to_number) if "weight" in df.columns else None
    out = out.dropna(subset=["company"])
    out = out[out["company"].astype(str).str.strip() != ""]

    if out["weight"].isna().all():
        market_value = out["market_value"].fillna(0.0).astype(float)
        total = float(market_value.sum())
        out["weight"] = (market_value / total) if total > 0 else 0.0
    return out


def _scan_label_value(df: pd.DataFrame, labels: set[str]) -> float | None:
    if df.empty:
        return None
    for row_idx in range(min(120, len(df))):
        row = df.iloc[row_idx].tolist()
        for col_idx in range(min(40, len(row))):
            if _normalize_label(row[col_idx]) not in labels:
                continue
            right = row[col_idx + 1] if col_idx + 1 < len(row) else None
            down = df.iloc[row_idx + 1, col_idx] if row_idx + 1 < len(df) and col_idx < len(df.columns) else None
            diag = (
                df.iloc[row_idx + 1, col_idx + 1]
                if row_idx + 1 < len(df) and col_idx + 1 < len(df.columns)
                else None
            )
            for candidate in (right, down, diag):
                number = _to_number(candidate)
                if number is not None:
                    return number
    return None


def parse_valuation_excel(path: Path, *, sheet: str = "") -> FundValuation:
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheets = [sheet] if sheet else list(xls.sheet_names)

    nav = None
    net_assets = None
    for sheet_name in sheets:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=object)
        nav = nav if nav is not None else _scan_label_value(df, VALUATION_NAV_LABELS)
        net_assets = net_assets if net_assets is not None else _scan_label_value(df, VALUATION_NET_ASSETS_LABELS)
        if nav is not None and net_assets is not None:
            break
    return FundValuation(nav=nav, net_assets=net_assets)


def _infer_attachment_kind(path: Path) -> str:
    name = str(path.name).lower()
    if any(keyword in name for keyword in VALUATION_FILENAME_KEYWORDS):
        return "valuation"
    if any(keyword in name for keyword in HOLDINGS_FILENAME_KEYWORDS):
        return "holdings"
    return "unknown"


def _looks_like_cross_broker_attachment(path: Path) -> bool:
    name = str(path.name).lower()
    return "statement" in name or "履约保障" in str(path.name)


def _resolve_effective_trade_date_for_paths(
    *,
    requested_trade_date: date,
    paths: Iterable[Path],
) -> tuple[date, list[Path], str | None]:
    file_list = list(paths)
    if not file_list:
        return requested_trade_date, [], None

    matched_paths: list[Path] = []
    available_dates: set[date] = set()
    date_cache: dict[Path, list[date]] = {}
    for path in file_list:
        parsed_dates = _extract_trade_dates_from_path(path)
        date_cache[path] = parsed_dates
        if not parsed_dates:
            continue
        available_dates.update(parsed_dates)
        if requested_trade_date in parsed_dates:
            matched_paths.append(path)

    if matched_paths:
        return requested_trade_date, matched_paths, None
    if not available_dates:
        return requested_trade_date, file_list, None

    fallback_date = max(available_dates)
    fallback_paths = [path for path in file_list if fallback_date in date_cache.get(path, [])]
    warning = (
        f"no attachment names matched trade date {requested_trade_date.isoformat()}, "
        f"fallback to latest file date {fallback_date.isoformat()}"
    )
    return fallback_date, (fallback_paths or file_list), warning


def _resolve_effective_product_trade_date_for_paths(
    *,
    requested_trade_date: date,
    paths: Iterable[Path],
) -> tuple[date, list[Path], str | None]:
    file_list = list(paths)
    if not file_list:
        return requested_trade_date, [], None

    requested_paths = [
        path for path in file_list if requested_trade_date in _extract_trade_dates_from_path(path)
    ]
    if requested_paths:
        return requested_trade_date, requested_paths, None

    available_dates = sorted(
        {trade_date for path in file_list for trade_date in _extract_trade_dates_from_path(path)}
    )
    if not available_dates:
        return requested_trade_date, file_list, None

    scored = [
        (score_product_inputs_for_date(file_list, candidate), candidate)
        for candidate in available_dates
    ]
    best_score, best_date = max(scored, key=lambda item: (item[0][0], item[0][1], item[1]))
    if best_score == (0, 0):
        return _resolve_effective_trade_date_for_paths(
            requested_trade_date=requested_trade_date,
            paths=file_list,
        )

    latest_date = max(available_dates)
    best_paths = [path for path in file_list if best_date in _extract_trade_dates_from_path(path)]
    if best_date == latest_date:
        return (
            best_date,
            best_paths,
            f"no attachment names matched trade date {requested_trade_date.isoformat()}, "
            f"fallback to latest product file date {best_date.isoformat()}",
        )
    return (
        best_date,
        best_paths,
        f"no attachment names matched trade date {requested_trade_date.isoformat()}, "
        f"fallback to most complete product file date {best_date.isoformat()} "
        f"(latest file date {latest_date.isoformat()}, "
        f"products={best_score[0]}, sources={best_score[1]})",
    )


def _product_email_completion_issues(
    *,
    results: list[dict[str, Any]],
    summary_path: Path | None,
    require_charts: bool,
    chart_paths: dict[str, Path],
) -> list[str]:
    issues: list[str] = []
    expected_names = [str(cfg.get("name", "")).strip() for cfg in PRODUCT_CONFIG]
    expected_names = [name for name in expected_names if name]
    by_name = {
        str(result.get("product_name", "")).strip(): result
        for result in results
        if str(result.get("product_name", "")).strip()
    }

    missing_products = [name for name in expected_names if name not in by_name]
    if missing_products:
        issues.append(f"missing product reports: {', '.join(missing_products)}")

    missing_nav = [name for name in expected_names if name in by_name and by_name[name].get("nav") is None]
    if missing_nav:
        issues.append(f"missing NAV: {', '.join(missing_nav)}")

    empty_holdings = [
        name
        for name in expected_names
        if name in by_name and int(by_name[name].get("total_holdings") or 0) <= 0
    ]
    if empty_holdings:
        issues.append(f"empty holdings: {', '.join(empty_holdings)}")

    missing_outputs: list[str] = []
    for name in expected_names:
        result = by_name.get(name)
        if result is None:
            continue
        out_xlsx = str(result.get("out_xlsx", "") or "").strip()
        if not out_xlsx or not Path(out_xlsx).exists():
            missing_outputs.append(name)
    if missing_outputs:
        issues.append(f"missing product Excel outputs: {', '.join(missing_outputs)}")

    if summary_path is None or not summary_path.exists():
        issues.append("missing summary Excel output")

    if require_charts:
        missing_charts = [
            name
            for name in expected_names
            if name in by_name and (
                chart_paths.get(name) is None or not chart_paths[name].exists()
            )
        ]
        if missing_charts:
            issues.append(f"missing charts: {', '.join(missing_charts)}")

    return issues


def _probe_standard_holdings_paths(paths: Iterable[Path]) -> list[Path]:
    candidates: list[Path] = []
    for path in paths:
        if _looks_like_cross_broker_attachment(path):
            continue
        try:
            parsed = parse_holdings_excel(path)
        except Exception:
            continue
        if parsed.empty:
            continue
        candidates.append(path)
    return candidates


def _build_cross_broker_company_report(
    *,
    trade_date: date,
    inbox_dir: Path,
    out_xlsx: Path,
) -> dict[str, Any]:
    with TemporaryDirectory(prefix="fund_portfolio_") as temp_dir:
        results = build_product_reports(
            trade_date=trade_date,
            inbox_dir=inbox_dir,
            out_dir=Path(temp_dir),
        )

    if not results:
        raise RuntimeError(
            f"cross-broker attachments detected, but no product holdings were built for {trade_date.isoformat()}"
        )

    frames: list[pd.DataFrame] = []
    product_names: list[str] = []
    source_files: set[str] = set()
    total_net_assets = 0.0
    has_total_net_assets = False
    for result in results:
        holdings_raw = result.get("holdings_raw")
        if holdings_raw is None or getattr(holdings_raw, "empty", True):
            continue

        frame = holdings_raw.copy()
        if "company" not in frame.columns or "market_value_cny" not in frame.columns:
            continue

        keep_columns = [column for column in ("company", "market_value_cny", "source_files") if column in frame.columns]
        frame = frame.loc[:, keep_columns]
        frame["company"] = frame["company"].astype(str).str.strip()
        frame["market_value_cny"] = pd.to_numeric(frame["market_value_cny"], errors="coerce")
        frame = frame.dropna(subset=["company", "market_value_cny"])
        frame = frame[frame["company"] != ""]
        if frame.empty:
            continue
        frames.append(frame)

        product_name = str(result.get("product_name", "") or "").strip()
        if product_name:
            product_names.append(product_name)

        nav_value = result.get("asset_nav")
        if nav_value is None:
            nav_value = result.get("nav")
        if nav_value is not None:
            total_net_assets += float(nav_value)
            has_total_net_assets = True

    if not frames:
        raise RuntimeError(
            f"cross-broker attachments detected, but no valid holdings rows were parsed for {trade_date.isoformat()}"
        )

    holdings = pd.concat(frames, ignore_index=True)
    if "source_files" in holdings.columns:
        for source_value in holdings["source_files"].dropna().astype(str):
            for item in source_value.split(","):
                name = item.strip()
                if name:
                    source_files.add(name)

    by_company = (
        holdings.groupby("company", as_index=False)
        .agg(market_value=("market_value_cny", "sum"))
        .sort_values("market_value", ascending=False, na_position="last")
        .reset_index(drop=True)
    )

    net_assets = total_net_assets if has_total_net_assets and total_net_assets > 0 else None
    if net_assets is not None:
        by_company["weight"] = by_company["market_value"].astype(float) / float(net_assets)
        by_company["weight_by_net_assets"] = by_company["weight"]
    else:
        by_company["weight"] = None
        by_company["weight_by_net_assets"] = None

    summary = pd.DataFrame(
        [
            {
                "trade_date": trade_date.isoformat(),
                "nav": net_assets,
                "net_assets": net_assets,
                "holdings_files": ",".join(sorted(source_files)),
                "valuation_files": "",
                "company_count": int(by_company["company"].nunique()),
                "mode": "cross_broker_aggregate",
                "product_count": len(results),
                "products": ",".join(product_names),
            }
        ]
    )

    _ensure_dir(out_xlsx.parent)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        by_company.to_excel(writer, sheet_name="by_company", index=False)

    return {
        "trade_date": trade_date.isoformat(),
        "out_xlsx": str(out_xlsx),
        "nav": summary.loc[0, "nav"],
        "net_assets": summary.loc[0, "net_assets"],
        "companies": int(summary.loc[0, "company_count"]),
        "mode": "cross_broker_aggregate",
        "product_count": int(summary.loc[0, "product_count"]),
    }


def build_portfolio_report(
    *,
    trade_date: date,
    valuation_paths: list[Path],
    holdings_paths: list[Path],
    out_xlsx: Path,
) -> dict[str, Any]:
    if not holdings_paths:
        raise ValueError("holdings_paths is empty")

    valuation = parse_valuation_excel(valuation_paths[0]) if valuation_paths else None

    holdings_frames: list[pd.DataFrame] = []
    for path in holdings_paths:
        df = parse_holdings_excel(path)
        df["source_file"] = path.name
        holdings_frames.append(df)
    holdings = pd.concat(holdings_frames, ignore_index=True)

    by_company = (
        holdings.groupby("company", as_index=False)
        .agg(
            market_value=("market_value", "sum"),
            weight=("weight", "sum"),
        )
        .sort_values("weight", ascending=False)
        .reset_index(drop=True)
    )

    net_assets = float(valuation.net_assets) if valuation and valuation.net_assets is not None else None
    if net_assets and net_assets > 0 and by_company["market_value"].notna().any():
        market_value = by_company["market_value"].fillna(0.0).astype(float)
        by_company["weight_by_net_assets"] = market_value / net_assets
    else:
        by_company["weight_by_net_assets"] = None

    summary = pd.DataFrame(
        [
            {
                "trade_date": trade_date.isoformat(),
                "nav": float(valuation.nav) if valuation and valuation.nav is not None else None,
                "net_assets": net_assets,
                "holdings_files": ",".join(path.name for path in holdings_paths),
                "valuation_files": ",".join(path.name for path in valuation_paths),
                "company_count": int(by_company["company"].nunique()),
            }
        ]
    )

    _ensure_dir(out_xlsx.parent)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        by_company.to_excel(writer, sheet_name="by_company", index=False)

    return {
        "trade_date": trade_date.isoformat(),
        "out_xlsx": str(out_xlsx),
        "nav": summary.loc[0, "nav"],
        "net_assets": summary.loc[0, "net_assets"],
        "companies": int(summary.loc[0, "company_count"]),
    }


def _resolve_imap_config_from_args(args: argparse.Namespace) -> ImapConfig:
    return ImapConfig(
        host=str(args.imap_host or get_env("IMAP_HOST", required=True) or ""),
        user=str(args.imap_user or get_env("IMAP_USER", required=True) or ""),
        password=str(args.imap_pass or get_env("IMAP_PASS", required=True) or ""),
        mailbox=str(args.imap_mailbox or get_env("IMAP_MAILBOX", "INBOX") or "INBOX"),
        port=int(args.imap_port or int(get_env("IMAP_PORT", "993") or "993")),
        use_ssl=bool(args.imap_ssl),
    )


def _resolve_state_file_arg(value: str, *, inbox_root: Path | None = None) -> Path | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    if raw.lower() == "default":
        return _default_mail_state_path(inbox_root=inbox_root)
    return Path(raw)


def _resolve_summary_xlsx_for_trade_date(
    *,
    requested_trade_date: date,
    summary_xlsx: str = "",
) -> tuple[date, Path, str | None]:
    raw = str(summary_xlsx or "").strip()
    if raw:
        return requested_trade_date, Path(raw), None

    report_root = _default_report_root()
    exact_path = report_root / requested_trade_date.isoformat() / f"fund_portfolio_summary_{requested_trade_date.isoformat()}.xlsx"
    if exact_path.exists():
        return requested_trade_date, exact_path, None

    candidates = sorted(report_root.glob("*/fund_portfolio_summary_*.xlsx"))
    effective_trade_date, matched_paths, warning = _resolve_effective_trade_date_for_paths(
        requested_trade_date=requested_trade_date,
        paths=candidates,
    )
    if matched_paths:
        return effective_trade_date, matched_paths[0], warning
    return requested_trade_date, exact_path, warning


def _build_portfolio_report_for_trade_date(
    *,
    trade_date: date,
    inbox_dir: Path | None = None,
    out_xlsx: Path | None = None,
) -> dict[str, Any]:
    resolved_inbox = Path(inbox_dir) if inbox_dir is not None else _default_inbox_dir(trade_date)
    if not resolved_inbox.exists():
        raise RuntimeError(f"inbox_dir not found: {resolved_inbox}")

    paths = sorted(path for path in resolved_inbox.glob("*.xls*") if path.is_file())
    effective_trade_date, effective_paths, warning = _resolve_effective_trade_date_for_paths(
        requested_trade_date=trade_date,
        paths=paths,
    )
    if warning:
        print(f"[WARN] {warning}")

    valuation_paths = [path for path in effective_paths if _infer_attachment_kind(path) == "valuation"]
    holdings_paths = [path for path in effective_paths if _infer_attachment_kind(path) == "holdings"]
    unknown_paths = [path for path in effective_paths if _infer_attachment_kind(path) == "unknown"]

    probed_holdings = _probe_standard_holdings_paths(unknown_paths)
    if probed_holdings:
        existing = set(holdings_paths)
        holdings_paths.extend([path for path in probed_holdings if path not in existing])

    if out_xlsx is not None:
        explicit_out = Path(out_xlsx)
        expected_name = f"portfolio_{trade_date.isoformat()}.xlsx"
        if explicit_out.name == expected_name and trade_date != effective_trade_date:
            resolved_out = explicit_out.with_name(f"portfolio_{effective_trade_date.isoformat()}.xlsx")
        else:
            resolved_out = explicit_out
    else:
        resolved_out = _default_portfolio_report_path(effective_trade_date)
    if not holdings_paths and any(_looks_like_cross_broker_attachment(path) for path in effective_paths):
        print(f"[INFO] fallback to cross-broker aggregate build for {effective_trade_date.isoformat()}")
        return _build_cross_broker_company_report(
            trade_date=effective_trade_date,
            inbox_dir=resolved_inbox,
            out_xlsx=resolved_out,
        )

    if not holdings_paths:
        raise RuntimeError(
            f"no standard holdings attachments found for {effective_trade_date.isoformat()} in {resolved_inbox}"
        )

    return build_portfolio_report(
        trade_date=effective_trade_date,
        valuation_paths=valuation_paths,
        holdings_paths=holdings_paths,
        out_xlsx=resolved_out,
    )


def _run_email_sync_for_trade_date(
    *,
    trade_date: date,
    args: argparse.Namespace,
    out_dir: Path | None = None,
    skip_processed: bool = False,
    state_file: Path | None = None,
) -> EmailSyncResult:
    load_env()
    resolved_out_dir = Path(out_dir) if out_dir is not None else (
        Path(args.out_dir) if str(getattr(args, "out_dir", "") or "").strip() else _default_inbox_dir(trade_date)
    )
    return _fetch_excel_attachments_via_imap_with_state(
        imap=_resolve_imap_config_from_args(args),
        target_date=trade_date,
        out_dir=resolved_out_dir,
        sender_allowlist=_parse_csv_arg(str(getattr(args, "sender_allowlist", "") or "")),
        subject_keywords=_parse_csv_arg(str(getattr(args, "subject_keywords", "") or "")),
        state_file=state_file,
        skip_processed=skip_processed,
        product_scope=bool(getattr(args, "product_scope", True)),
        scope_from_products=bool(getattr(args, "scope_from_products", True)),
        # full-access 邮箱适配: 只扫最新若干封 + 客户端按收件日筛，避免拉全箱历史/掉线。
        # 2000 封约覆盖最近 20 天收件(~100/日)，足够日常增量与 10 天补发窗口。
        recent_scan_limit=int(getattr(args, "recent_scan_limit", 2000) or 2000),
        enforce_received_date=bool(getattr(args, "enforce_received_date", True)),
    )


def _infer_broker_from_name(name: str) -> str:
    """从附件文件名 best-effort 推断券商。"""
    low = name.lower()
    if "statement" in low or "履约保障" in name or "citic" in low or "中信" in name:
        return "citic"
    if "cicc" in low or "中金" in name:
        return "cicc"
    if "swhysc" in low or "申万" in name or "宏源" in name:
        return "swhysc"
    return "unknown"


def _infer_product_code_from_name(name: str) -> str | None:
    """从附件文件名 best-effort 反查产品代码（命中 cicc/citic 账户号即返回）。"""
    from fundadmin.clients.config import NAME_TO_PRODCODE

    low = name.lower()
    for cfg in PRODUCT_CONFIG:
        cicc = [str(t) for t in (cfg.get("cicc_codes") or [])]
        citic = [str(t) for t in (cfg.get("citic_codes") or [])]
        prodcode = cicc[0] if cicc else NAME_TO_PRODCODE.get(cfg.get("name", ""))
        for tok in cicc + citic:
            if tok and tok.lower() in low:
                return prodcode
    return None


def _persist_attachments_raw(
    files: list[Path],
    *,
    effective_trade_date: date,
    inbox_dir: Path,
) -> None:
    """原始无损落地层：每个附件按内容 sha256 去重，逐 sheet 逐行存 JSON。

    入库失败仅告警，不影响报表/邮件主流程。
    """
    from fundadmin.clients.schema import init_db
    from fundadmin.clients.store import insert_attachment, upsert_raw_sheet_rows
    from fundadmin.portfolio.parsers.common import read_csv_robust

    try:
        init_db()
    except Exception:
        logger.exception("raw-layer: init_db failed; skip ingest")
        return

    ingested = 0
    skipped = 0
    for path in files:
        try:
            data = path.read_bytes()
            sha = hashlib.sha256(data).hexdigest()
            suffix = path.suffix.lower()

            sheets: list[tuple[str, pd.DataFrame]] = []
            if suffix in (".xlsx", ".xls"):
                book = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
                sheets = list(book.items())
            elif suffix == ".csv":
                sheets = [("csv", read_csv_robust(path))]

            total_rows = sum(len(df) for _, df in sheets)
            meta = {
                "sha256": sha,
                "file_name": path.name,
                "file_suffix": suffix,
                "broker": _infer_broker_from_name(path.name),
                "product_code": _infer_product_code_from_name(path.name),
                "as_of_date": effective_trade_date.isoformat(),
                "attachment_type": _infer_attachment_kind(path),
                "sheet_count": len(sheets),
                "row_count": int(total_rows),
                "inbox_dir": str(inbox_dir),
            }
            ingest_id, is_new = insert_attachment(meta)
            if not is_new:
                skipped += 1
                continue

            rows: list[dict[str, Any]] = []
            for s_idx, (s_name, df) in enumerate(sheets):
                for r_idx, (_, row) in enumerate(df.iterrows()):
                    cells = [None if pd.isna(v) else v for v in row.tolist()]
                    rows.append(
                        {
                            "ingest_id": ingest_id,
                            "sheet_index": s_idx,
                            "sheet_name": str(s_name),
                            "row_index": r_idx,
                            "cells_json": json.dumps(cells, ensure_ascii=False, default=str),
                        }
                    )
            if rows:
                upsert_raw_sheet_rows(pd.DataFrame(rows))
            ingested += 1
        except Exception:
            logger.exception("raw-layer ingest failed for %s", path)

    print(f"[OK] raw-layer ingest: {ingested} new attachment(s), {skipped} duplicate(s) skipped")


def _persist_curated_layer(
    results: list[dict[str, Any]],
    *,
    effective_trade_date: date,
) -> None:
    """核心结构层：写 product_valuation + fund_positions（分券商）。入库失败仅告警。

    fund_positions 按 (产品, 标的, 券商) 逐行落地：同一标的若同时在中金、中信持有，
    各券商单独成行（broker 真实填充），不再跨券商折叠。写入前先删除该产品该估值日的
    旧快照，保证重跑幂等、不与旧 broker='' 折叠行重复计数。
    """
    from fundadmin.clients.config import NAME_TO_PRODCODE
    from fundadmin.clients.store import (
        delete_positions,
        upsert_positions,
        upsert_product_valuation,
    )

    as_of = effective_trade_date.isoformat()
    val_rows: list[dict[str, Any]] = []
    pos_by_pcode: dict[str, list[pd.DataFrame]] = {}

    for r in results:
        pname = r.get("product_name", "")
        pcode = NAME_TO_PRODCODE.get(pname)
        if not pcode:
            # 未在映射中的产品（如 沐泽1号 SQJ420）暂不入结构层；原始层已无损保留。
            logger.warning("curated-layer: no product_code for %s, skipped structured tables", pname)
            continue

        val_rows.append(
            {
                "as_of_date": as_of,
                "product_code": pcode,
                "product_name": pname,
                "unit_nav": r.get("unit_nav"),
                "asset_nav": r.get("asset_nav"),
                "nav_for_weight": r.get("nav"),
                "total_holdings": r.get("total_holdings"),
                "total_market_value_cny": r.get("total_market_value_cny"),
                "ingest_id": None,
            }
        )

        # 优先用分券商明细；旧 payload 无该键时回退到合并视图（broker 落为 ''）。
        positions = r.get("holdings_by_broker")
        if positions is None or getattr(positions, "empty", True):
            positions = r.get("holdings_raw")
        if positions is not None and not positions.empty:
            h = positions.copy()
            h["as_of_date"] = as_of
            h["product_code"] = pcode
            h["product_name"] = pname
            h = h.rename(
                columns={
                    "company": "instrument_name",
                    "shares": "quantity",
                    "source_files": "source_files",
                }
            )
            pos_by_pcode.setdefault(pcode, []).append(h)

    try:
        n_val = upsert_product_valuation(pd.DataFrame(val_rows)) if val_rows else 0
        n_pos = 0
        for pcode, frames in pos_by_pcode.items():
            # 先清旧快照再写新分券商行：重跑幂等，并清理已清仓的残留标的。
            delete_positions(pcode, as_of)
            n_pos += upsert_positions(pd.concat(frames, ignore_index=True))
        print(f"[OK] curated-layer: {n_val} valuation row(s), {n_pos} position row(s, per-broker)")
    except Exception:
        logger.exception("curated-layer persistence failed")


def _build_tx_product_lookup() -> dict[str, tuple[str, str]]:
    """构造"代码 token -> (product_code, product_name)"映射，用于从文件名反查产品。

    同时收录 CITIC 账户号（如 104902）与 CICC 代码（如 SCD704），统一映射到
    curated 层使用的规范 product_code（NAME_TO_PRODCODE）。
    """
    from fundadmin.clients.config import NAME_TO_PRODCODE

    lut: dict[str, tuple[str, str]] = {}
    for cfg in PRODUCT_CONFIG:
        pname = cfg.get("name", "")
        pcode = NAME_TO_PRODCODE.get(pname)
        if not pcode:
            continue
        for tok in (cfg.get("citic_codes") or []) + (cfg.get("cicc_codes") or []):
            if tok:
                lut[str(tok)] = (pcode, pname)
    return lut


def _persist_transactions(
    files: list[Path],
    *,
    effective_trade_date: date,
) -> None:
    """成交流水层：解析 CICC"当日交易" + CITIC"Transaction"全量入库（按 occ 去重）。

    范围为全历史成交：
    - CITIC Statement 的 Transaction sheet 为"当日"逐笔成交，故需遍历 inbox 中
      所有日期的 Statement 文件（并非仅目标日），方能累积全历史。
    - CICC"当日交易"sheet 为全历史成交，单份报告即含全量；occ 负责跨文件折叠重复。
    跨文件/跨快照去重靠主键 + occ；入库失败仅告警，不阻断主流程。
    """
    from fundadmin.clients.store import upsert_transactions
    from fundadmin.portfolio.parsers.trades import (
        parse_cicc_trades,
        parse_citic_transactions,
    )

    if not files:
        return

    lut = _build_tx_product_lookup()
    tx_frames: list[pd.DataFrame] = []

    for p in files:
        name = p.name
        suffix = p.suffix.lower()
        # 反查产品：文件名中命中任一代码 token。
        match = next((v for tok, v in lut.items() if tok in name), None)
        if match is None:
            continue
        pcode, pname = match

        if "Statement" in name and suffix in {".xlsx", ".xlsm", ".xls"}:
            # CITIC 履约保障报告：Transaction sheet（当日逐笔成交）。
            try:
                tx = parse_citic_transactions(p)
            except Exception:
                logger.debug("parse_citic_transactions failed for %s", p, exc_info=True)
                continue
            broker = "citic"
        elif suffix == ".xlsx" and "Statement" not in name:
            # CICC 估值报告附件：当日交易 sheet（全历史成交）。无该 sheet 返回空表。
            try:
                tx = parse_cicc_trades(p)
            except Exception:
                logger.debug("parse_cicc_trades failed for %s", p, exc_info=True)
                continue
            broker = "cicc"
        else:
            continue

        if tx is not None and not tx.empty:
            tx = tx.copy()
            tx["broker"] = broker
            tx["product_code"] = pcode
            tx["product_name"] = pname
            tx_frames.append(tx)

    if not tx_frames:
        print("[OK] transactions: 0 row(s) (no trade records found)")
        return

    try:
        n_tx = upsert_transactions(pd.concat(tx_frames, ignore_index=True))
        print(f"[OK] transactions: {n_tx} row(s) upserted")
    except Exception:
        logger.exception("transactions persistence failed")


def _build_product_reports_for_trade_date(
    *,
    trade_date: date,
    inbox_dir: Path | None = None,
    report_root: Path | None = None,
    out_dir: Path | None = None,
    with_charts: bool = False,
    with_email: bool = False,
    notify_clients: bool = False,
    email_to: str = "",
    smtp_host: str = "",
    smtp_port: int = 0,
    smtp_user: str = "",
    smtp_pass: str = "",
    smtp_from: str = "",
    summary_alias_xlsx: Path | None = None,
) -> dict[str, Any]:
    resolved_inbox = Path(inbox_dir) if inbox_dir is not None else _default_inbox_dir(trade_date)
    files = [path for path in resolved_inbox.iterdir() if path.is_file()] if resolved_inbox.exists() else []
    effective_trade_date, _, warning = _resolve_effective_product_trade_date_for_paths(
        requested_trade_date=trade_date,
        paths=files,
    )
    if warning:
        print(f"[WARN] {warning}")

    # 原始无损落地层：把本次 inbox 的所有附件按 sha256 去重入库（失败仅告警）。
    if files:
        _persist_attachments_raw(
            files,
            effective_trade_date=effective_trade_date,
            inbox_dir=resolved_inbox,
        )

    resolved_report_root = Path(report_root) if report_root is not None else _default_report_root()
    resolved_out_dir = Path(out_dir) if out_dir is not None else (
        resolved_report_root / effective_trade_date.isoformat()
    )
    try:
        results = build_product_reports(
            trade_date=effective_trade_date,
            inbox_dir=resolved_inbox,
            out_dir=resolved_out_dir,
        )
    except PermissionError as exc:
        raise RuntimeError(
            f"cannot overwrite product report output, file may be open in Excel: {exc}"
        ) from exc
    print(f"[OK] {len(results)} product reports built into: {resolved_out_dir}")
    for result in results:
        print(f"  {result['product_name']}: {result['out_xlsx']}")

    # 核心结构层：从 build 结果写 product_valuation + fund_positions（失败仅告警）。
    _persist_curated_layer(results, effective_trade_date=effective_trade_date)

    # 成交流水层：解析 CICC/CITIC 成交全量入库（按 occ 去重；失败仅告警）。
    _persist_transactions(files, effective_trade_date=effective_trade_date)

    chart_paths: dict[str, Path] = {}
    if with_charts:
        charts_dir = resolved_out_dir / "charts"
        for result in results:
            pname = result.get("product_name", "")
            holdings_raw = result.get("holdings_raw")
            if holdings_raw is None or holdings_raw.empty:
                print(f"[WARN] {pname}: no holdings data for chart generation")
                continue
            chart_path = charts_dir / f"{pname}_holdings_pie.png"
            try:
                generate_portfolio_pie_chart(
                    holdings_raw,
                    product_name=pname,
                    trade_date=effective_trade_date,
                    nav=result.get("nav"),
                    total_holdings=result.get("total_holdings", 0),
                    out_path=chart_path,
                )
                chart_paths[pname] = chart_path
                print(f"[OK] {pname}: chart saved to {chart_path}")
            except Exception:
                logger.exception("%s: chart generation failed", pname)

    summary_path = resolved_out_dir / f"fund_portfolio_summary_{effective_trade_date.isoformat()}.xlsx"
    try:
        build_summary_excel(results, trade_date=effective_trade_date, out_path=summary_path)
        print(f"[OK] summary excel saved to: {summary_path}")
    except PermissionError as exc:
        raise RuntimeError(
            f"cannot overwrite summary output, file may be open in Excel: {exc}"
        ) from exc
    except Exception:
        logger.exception("summary excel generation failed")
        summary_path = None

    legacy_summary_path: Path | None = None
    if summary_alias_xlsx is not None and summary_path is not None and summary_path.exists():
        try:
            legacy_summary_path = Path(summary_alias_xlsx)
            if legacy_summary_path != summary_path:
                _ensure_dir(legacy_summary_path.parent)
                legacy_summary_path.write_bytes(summary_path.read_bytes())
                print(
                    "[WARN] build is now an alias of build-products; "
                    f"summary copied to legacy path: {legacy_summary_path}"
                )
        except PermissionError as exc:
            raise RuntimeError(
                f"cannot write legacy summary alias, file may be open in Excel: {exc}"
            ) from exc

    email_sent = False
    email_skip_reason = ""
    client_notify_sent = 0
    client_notify_skip_reason = ""
    # 两封邮件（内部持仓汇总 + 客户净值通知）共用同一完整性门槛：
    # 仅当成功构建且数据完整（_product_email_completion_issues 为空）时才发送。
    if with_email or notify_clients:
        completion_issues = _product_email_completion_issues(
            results=results,
            summary_path=summary_path,
            require_charts=with_charts,
            chart_paths=chart_paths,
        )
        if completion_issues:
            reason = "; ".join(completion_issues)
            if with_email:
                email_skip_reason = reason
                print(f"[WARN] matrix email skipped: {reason}")
            if notify_clients:
                client_notify_skip_reason = reason
                print(f"[WARN] client NAV notify skipped: {reason}")
        else:
            load_env()

            # ---- 内部持仓汇总邮件（EMAIL_TO，QQ 邮箱 SMTP）----
            if with_email:
                to_addrs = [x.strip() for x in str(email_to or "").split(",") if x.strip()]
                if not to_addrs:
                    env_to = get_env("EMAIL_TO", default="")
                    to_addrs = [x.strip() for x in str(env_to).split(",") if x.strip()]
                if not to_addrs:
                    raise RuntimeError("--with-email requires --email-to or EMAIL_TO")

                resolved_smtp_host = str(smtp_host or get_env("SMTP_HOST", required=True) or "")
                resolved_smtp_port = int(smtp_port or int(get_env("SMTP_PORT", "465") or "465"))
                resolved_smtp_user = str(smtp_user or get_env("SMTP_USER", required=True) or "")
                resolved_smtp_pass = str(smtp_pass or get_env("SMTP_PASS", required=True) or "")
                resolved_smtp_from = str(
                    smtp_from or get_env("EMAIL_FROM", default="") or resolved_smtp_user
                )

                if not resolved_smtp_host or not resolved_smtp_user or not resolved_smtp_pass:
                    raise RuntimeError("SMTP config incomplete: require SMTP_HOST, SMTP_USER, SMTP_PASS")

                smtp = SmtpConfig(
                    host=resolved_smtp_host,
                    port=resolved_smtp_port,
                    user=resolved_smtp_user,
                    password=resolved_smtp_pass,
                    from_addr=resolved_smtp_from,
                )

                attachments: dict[str, bytes] = {}
                if summary_path is not None and summary_path.exists():
                    attachments[summary_path.name] = summary_path.read_bytes()

                send_matrix_email(
                    results,
                    trade_date=effective_trade_date,
                    chart_paths=chart_paths,
                    smtp_config=smtp,
                    to_addrs=to_addrs,
                    attachments=attachments if attachments else None,
                )
                excluded = {"沐泽1号"}
                sent_count = sum(1 for result in results if result.get("product_name", "") not in excluded)
                email_sent = True
                print(f"[OK] matrix email ({sent_count} products) sent to: {', '.join(to_addrs)}")

                # ---- 分券商组合持仓汇总邮件（同收件人/同 SMTP；失败仅告警，不影响主流程）----
                try:
                    bb_sent = send_by_broker_summary_email(
                        results,
                        trade_date=effective_trade_date,
                        smtp_config=smtp,
                        to_addrs=to_addrs,
                    )
                    if bb_sent:
                        print(f"[OK] by-broker summary email sent to: {', '.join(to_addrs)}")
                    else:
                        print("[INFO] by-broker summary email skipped: no per-broker data")
                except Exception:
                    logger.exception("by-broker summary email failed")

            # ---- 客户净值通知（clients 表，企业邮箱 SMTP xuekun@hysttz.com）----
            if notify_clients:
                nc_user = str(get_env("IMAP_USER", default="") or "")
                nc_pass = str(get_env("IMAP_PASS", default="") or "")
                if not nc_user or not nc_pass:
                    client_notify_skip_reason = "SMTP config incomplete (need IMAP_USER/IMAP_PASS)"
                    print(f"[WARN] client NAV notify skipped: {client_notify_skip_reason}")
                elif summary_path is None or not summary_path.exists():
                    client_notify_skip_reason = "summary excel missing"
                    print(f"[WARN] client NAV notify skipped: {client_notify_skip_reason}")
                else:
                    nc_smtp = SmtpConfig(
                        host="smtp.exmail.qq.com",
                        port=465,
                        user=nc_user,
                        password=nc_pass,
                        from_addr=nc_user,
                    )
                    stats = send_client_nav_emails(
                        trade_date=effective_trade_date,
                        summary_xlsx=summary_path,
                        smtp_config=nc_smtp,
                    )
                    client_notify_sent = int(stats.get("sent", 0))
                    print(
                        f"[OK] client NAV notify: {stats['sent']} sent, "
                        f"{stats['skipped']} skipped, {stats.get('total', 0)} total"
                    )

    payload = {
        "trade_date": effective_trade_date.isoformat(),
        "out_dir": str(resolved_out_dir),
        "summary_xlsx": str(summary_path) if summary_path is not None else "",
        "product_count": len(results),
        "chart_count": len(chart_paths),
        "legacy_summary_xlsx": str(legacy_summary_path) if legacy_summary_path is not None else "",
    }
    if with_email:
        payload["email_sent"] = email_sent
        payload["email_skip_reason"] = email_skip_reason
    if notify_clients:
        payload["client_notify_sent"] = client_notify_sent
        payload["client_notify_skip_reason"] = client_notify_skip_reason
    return payload


def _cmd_email_sync(args: argparse.Namespace) -> int:
    load_env()
    trade_date = _resolve_trade_date_arg(args.trade_date)
    if bool(getattr(args, "print_search", False)):
        # 干跑：仅打印将要发给 IMAP 服务端的 SEARCH 表达式，不联网
        sender_tokens = _parse_csv_arg(str(getattr(args, "sender_allowlist", "") or ""))
        subject_kw = _parse_csv_arg(str(getattr(args, "subject_keywords", "") or ""))
        criteria = build_imap_search_criteria(
            since=trade_date,
            before=trade_date + timedelta(days=1),
            sender_allowlist=sorted({t.strip().lower() for t in sender_tokens if t.strip()}),
            subject_keywords=[t.strip().lower() for t in subject_kw if t.strip()],
            scope_from_products=bool(getattr(args, "scope_from_products", True)),
        )
        print(f"[INFO] trade_date={trade_date.isoformat()}")
        print(f"[INFO] scope_from_products={bool(getattr(args, 'scope_from_products', True))}")
        print(f"[INFO] IMAP SEARCH criteria:\n{criteria}")
        return 0
    state_value = str(getattr(args, "state_file", "") or "")
    if bool(getattr(args, "skip_processed", False)) and not state_value.strip():
        state_value = "default"
    state_file = _resolve_state_file_arg(
        state_value,
        inbox_root=_default_inbox_root(),
    )
    result = _run_email_sync_for_trade_date(
        trade_date=trade_date,
        args=args,
        skip_processed=bool(getattr(args, "skip_processed", False)),
        state_file=state_file,
    )
    print(f"[OK] saved {len(result.saved_paths)} excel attachments into: {result.out_dir}")
    print(
        f"matched_messages: {result.matched_messages}, "
        f"processed_messages: {result.processed_messages}, "
        f"skipped_messages: {result.skipped_messages}"
    )
    if result.state_file is not None:
        print(f"state_file: {result.state_file}")
    for path in result.saved_paths:
        print(str(path))
    return 0


def _cmd_prune_inbox(args: argparse.Namespace) -> int:
    """手动清理 fund_inbox 下旧日期子目录（默认 dry-run 友好）。"""
    inbox_root = Path(getattr(args, "inbox_root", "") or "") if str(getattr(args, "inbox_root", "") or "").strip() else _default_inbox_root()
    keep_last = int(getattr(args, "keep_last", 30) or 30)
    dry_run = bool(getattr(args, "dry_run", False))

    try:
        result = prune_inbox(inbox_root=inbox_root, keep_last=keep_last, dry_run=dry_run)
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 2

    action = "WOULD DELETE" if dry_run else "DELETED"
    print(f"[INFO] inbox_root: {result.inbox_root}")
    print(f"[INFO] keep_last:  {result.keep_last} (dry_run={dry_run})")
    print(f"[INFO] kept:       {len(result.kept_dirs)} date dirs")
    for path in result.kept_dirs:
        print(f"  KEEP   {path.name}")
    print(f"[INFO] {action.lower()}: {len(result.pruned_dirs)} date dirs, "
          f"~{format_bytes(result.bytes_pruned)} freed")
    for path in result.pruned_dirs:
        print(f"  {action} {path}")
    if result.skipped_dirs:
        print(f"[INFO] non-date entries skipped: {len(result.skipped_dirs)} (untouched)")
        for path in result.skipped_dirs:
            print(f"  SKIP   {path.name}")
    if result.failures:
        print(f"[WARN] {len(result.failures)} failures while deleting:")
        for path, err in result.failures:
            print(f"  FAIL   {path}: {err}")
        return 1
    return 0


def _cmd_build(args: argparse.Namespace) -> int:
    print("[WARN] build is now a compatibility alias of build-products. Prefer build-products.")
    trade_date = _resolve_trade_date_arg(args.trade_date)
    try:
        payload = _build_product_reports_for_trade_date(
            trade_date=trade_date,
            inbox_dir=Path(args.inbox_dir) if str(args.inbox_dir or "").strip() else None,
            out_dir=Path(getattr(args, "out_dir", "")) if str(getattr(args, "out_dir", "") or "").strip() else None,
            with_charts=bool(getattr(args, "with_charts", False)),
            with_email=bool(getattr(args, "with_email", False)),
            email_to=str(getattr(args, "email_to", "") or ""),
            smtp_host=str(getattr(args, "smtp_host", "") or ""),
            smtp_port=int(getattr(args, "smtp_port", 0) or 0),
            smtp_user=str(getattr(args, "smtp_user", "") or ""),
            smtp_pass=str(getattr(args, "smtp_pass", "") or ""),
            smtp_from=str(getattr(args, "smtp_from", "") or ""),
            summary_alias_xlsx=Path(args.out_xlsx) if str(args.out_xlsx or "").strip() else None,
        )
    except Exception:
        logger.exception("product reports failed")
        return 1
    print("[OK] product reports built")
    for key, value in payload.items():
        print(f"{key}: {value}")
    return 0


def _collect_inbox_files_by_trade_date(
    *, inbox_root: Path, asof_date: date, scan_days: int
) -> dict[date, list[Path]]:
    """扫描近 scan_days 天的收件日目录，按文件名嵌入的交易日重新分组。

    券商报告 T+1/T+2 才到邮箱，同一交易日的文件常散落在多个收件日目录，
    这里跨目录汇集，键为文件名解析出的交易日（仅保留不晚于 asof 的日期）。
    """
    grouped: dict[date, list[Path]] = {}
    if not inbox_root.exists():
        return grouped
    earliest = asof_date - timedelta(days=max(0, int(scan_days)))
    for child in sorted(inbox_root.iterdir()):
        if not child.is_dir() or child.name.startswith("_"):
            continue
        try:
            received = _parse_ymd(child.name)
        except ValueError:
            continue
        if received < earliest or received > asof_date:
            continue
        for path in sorted(child.iterdir()):
            if not path.is_file():
                continue
            for embedded in _extract_trade_dates_from_path(path):
                if embedded > asof_date or embedded < earliest:
                    continue
                grouped.setdefault(embedded, []).append(path)
    return grouped


def _stage_trade_date_files(
    *, inbox_root: Path, trade_date: date, files: list[Path]
) -> Path:
    """把某交易日的文件（可能跨多个收件目录）复制进单交易日暂存目录。

    单交易日目录可让原始无损层与成交流水层避免被同目录的其它交易日文件污染。
    同名文件按内容去重，只保留首个。
    """
    staged_root = inbox_root / "_staged" / trade_date.isoformat()
    if staged_root.exists():
        shutil.rmtree(staged_root)
    _ensure_dir(staged_root)
    seen_names: set[str] = set()
    for src in files:
        if not src.is_file() or src.name in seen_names:
            continue
        seen_names.add(src.name)
        shutil.copy2(src, staged_root / src.name)
    return staged_root


def _cmd_sync_latest(args: argparse.Namespace) -> int:
    load_env()
    asof_date = _parse_optional_ymd(str(getattr(args, "asof", "") or "")) or date.today()
    lookback = max(1, int(getattr(args, "lookback", DEFAULT_SYNC_LOOKBACK) or DEFAULT_SYNC_LOOKBACK))
    resolved_dates, warning = _resolve_latest_trade_dates(asof_date=asof_date, lookback=lookback)
    if warning:
        print(f"[WARN] {warning}")
    if not resolved_dates:
        raise RuntimeError("cannot resolve latest trade dates")

    inbox_root = Path(args.inbox_root) if str(getattr(args, "inbox_root", "") or "").strip() else _default_inbox_root()
    report_root = Path(args.report_root) if str(getattr(args, "report_root", "") or "").strip() else _default_report_root()
    state_file = _resolve_state_file_arg(
        str(getattr(args, "state_file", "default") or "default"),
        inbox_root=inbox_root,
    )

    # ---- 下载段：按收件日窗口拉取附件进 inbox/<收件日>/（IMAP 仅能按收件日检索）----
    total_saved = 0
    total_skipped = 0
    download_failures = 0
    for received_date in resolved_dates:
        out_dir = inbox_root / received_date.isoformat()
        # 单个收件日的 IMAP 失败（如服务端断连）不应拖垮整次同步：
        # 记录告警后继续，使发布段仍能用已下载到本地的文件发出已补齐的交易日。
        try:
            result = _run_email_sync_for_trade_date(
                trade_date=received_date,
                args=args,
                out_dir=out_dir,
                skip_processed=bool(getattr(args, "skip_processed", True)),
                state_file=state_file,
            )
        except Exception:
            download_failures += 1
            logger.exception("email sync failed for received_date %s", received_date.isoformat())
            print(f"[WARN] sync failed for {received_date.isoformat()}; continuing with already-downloaded files")
            continue
        total_saved += len(result.saved_paths)
        total_skipped += int(result.skipped_messages)
        print(
            f"[OK] synced {received_date.isoformat()} "
            f"(saved={len(result.saved_paths)}, matched={result.matched_messages}, skipped={result.skipped_messages})"
        )

    # ---- 发布段：按文件名里的交易日重新分组，逐个交易日构建并发送 ----
    # 券商报告 T+1/T+2 才到邮箱，一个收件日目录混着多个交易日；按交易日分组并跨目录汇集，
    # 保证较旧但已补齐的交易日也能发出。已发渠道由 published_state 去重，不齐的留待次日补发。
    built_reports = 0
    matrix_sent = 0
    clients_sent = 0
    if bool(getattr(args, "build", True)):
        with_email = bool(getattr(args, "with_email", False))
        notify_clients = bool(getattr(args, "notify_clients", False))
        republish = bool(getattr(args, "republish", False))
        publish_state_path = _default_publish_state_path(inbox_root)
        publish_state = _load_publish_state(publish_state_path)

        grouped = _collect_inbox_files_by_trade_date(
            inbox_root=inbox_root, asof_date=asof_date, scan_days=PUBLISH_SCAN_DAYS
        )
        emailing = with_email or notify_clients
        for trade_date in sorted(grouped):
            want_matrix = with_email and (
                republish or not _publish_channel_done(publish_state, trade_date, "matrix")
            )
            want_clients = notify_clients and (
                republish or not _publish_channel_done(publish_state, trade_date, "clients")
            )
            if emailing:
                # 发邮件模式：该交易日所有请求渠道都已发布则跳过，避免重复发送/重复解析。
                if not want_matrix and not want_clients:
                    continue
            else:
                # 纯同步模式：仅为尚无报表的交易日补建报表，已存在的不重复解析。
                if (report_root / trade_date.isoformat()).exists():
                    continue

            staged_dir = _stage_trade_date_files(
                inbox_root=inbox_root, trade_date=trade_date, files=grouped[trade_date]
            )
            try:
                payload = _build_product_reports_for_trade_date(
                    trade_date=trade_date,
                    inbox_dir=staged_dir,
                    report_root=report_root,
                    with_email=want_matrix,
                    notify_clients=want_clients,
                    email_to=str(getattr(args, "email_to", "") or ""),
                )
            except Exception:
                logger.exception("build skipped for %s", trade_date.isoformat())
                continue
            built_reports += 1
            print(
                f"[OK] built product reports for {payload['trade_date']}: "
                f"{payload['summary_xlsx'] or payload['out_dir']}"
            )
            if want_matrix and payload.get("email_sent"):
                _mark_publish_channel(publish_state, trade_date, "matrix")
                matrix_sent += 1
            if want_clients and int(payload.get("client_notify_sent") or 0) > 0:
                _mark_publish_channel(publish_state, trade_date, "clients")
                clients_sent += 1
            _save_publish_state(publish_state_path, publish_state)

    print(
        f"[OK] sync-latest finished: received_dates={len(resolved_dates)}, "
        f"download_failures={download_failures}, "
        f"saved_attachments={total_saved}, skipped_messages={total_skipped}, "
        f"built_reports={built_reports}, matrix_sent={matrix_sent}, clients_sent={clients_sent}"
    )
    if state_file is not None:
        print(f"state_file: {state_file}")
    # 仅当所有收件日下载都失败（彻底没拉到任何数据）才以非零退出，提示运维介入；
    # 部分失败但仍完成构建/发送时视为成功，避免 launchd 误报。
    if download_failures and download_failures == len(resolved_dates):
        return 1
    return 0


def _split_paths(value: str) -> list[Path]:
    return [Path(p.strip()) for p in str(value or "").split(",") if p.strip()]


def _cmd_build_cross_broker(args: argparse.Namespace) -> int:
    trade_date = _resolve_trade_date_arg(args.trade_date)

    inputs = CrossBrokerInput(
        valuation_path=Path(args.valuation_path) if args.valuation_path else None,
        cicc_paths=_split_paths(args.cicc_paths),
        citic_usd_underlying_paths=_split_paths(args.citic_usd_underlying_paths),
        citic_usd_balance_paths=_split_paths(args.citic_usd_balance_paths),
        citic_hkd_underlying_paths=_split_paths(args.citic_hkd_underlying_paths),
        citic_hkd_balance_paths=_split_paths(args.citic_hkd_balance_paths),
        swhysc_valuation_paths=_split_paths(args.swhysc_valuation_paths),
    )

    out_xlsx = Path(
        args.out_xlsx
        or (_default_report_root() / f"cross_broker_{trade_date.isoformat()}.xlsx")
    )
    payload = build_cross_broker_report(
        trade_date=trade_date,
        inputs=inputs,
        out_xlsx=out_xlsx,
    )
    print("[OK] cross-broker portfolio built")
    for key, value in payload.items():
        print(f"{key}: {value}")
    return 0


def _cmd_build_products(args: argparse.Namespace) -> int:
    trade_date = _resolve_trade_date_arg(args.trade_date)
    try:
        payload = _build_product_reports_for_trade_date(
            trade_date=trade_date,
            inbox_dir=Path(args.inbox_dir) if str(args.inbox_dir or "").strip() else None,
            out_dir=Path(args.out_dir) if str(args.out_dir or "").strip() else None,
            with_charts=bool(args.with_charts),
            with_email=bool(args.with_email),
            notify_clients=bool(getattr(args, "notify_clients", False)),
            email_to=str(args.email_to or ""),
            smtp_host=str(args.smtp_host or ""),
            smtp_port=int(args.smtp_port or 0),
            smtp_user=str(args.smtp_user or ""),
            smtp_pass=str(args.smtp_pass or ""),
            smtp_from=str(args.smtp_from or ""),
        )
    except Exception:
        logger.exception("product report build failed")
        return 1
    for key, value in payload.items():
        print(f"{key}: {value}")
    return 0


def _cmd_notify_clients(args: argparse.Namespace) -> int:
    """向客户发送分产品净值通知邮件。"""
    trade_date = _resolve_trade_date_arg(args.trade_date)
    effective_trade_date, summary_xlsx, warning = _resolve_summary_xlsx_for_trade_date(
        requested_trade_date=trade_date,
        summary_xlsx=str(args.summary_xlsx or ""),
    )
    if warning:
        print(f"[WARN] {warning}")
    if not summary_xlsx.exists():
        print(f"[ERROR] Summary Excel not found: {summary_xlsx}")
        return 1

    load_env()
    # 客户通知邮件使用企业邮箱 SMTP（xuekun@hysttz.com）
    # 默认使用腾讯企业邮箱 SMTP，与 IMAP 同账号，避免与 .env 中 QQ 邮箱 SMTP 冲突
    smtp_host = str(args.smtp_host or "smtp.exmail.qq.com")
    smtp_port = int(args.smtp_port or 465)
    smtp_user = str(args.smtp_user or get_env("IMAP_USER", default="") or "")
    smtp_pass = str(args.smtp_pass or get_env("IMAP_PASS", default="") or "")
    smtp_from = str(args.smtp_from or get_env("IMAP_USER", default="") or smtp_user)

    if not smtp_host or not smtp_user or not smtp_pass:
        print("[ERROR] SMTP 配置不完整：需要 IMAP_USER/IMAP_PASS 或显式传入 --smtp-user/--smtp-pass")
        return 1

    print(f"[INFO] SMTP: {smtp_user} via {smtp_host}:{smtp_port}, From: {smtp_from}")

    smtp = SmtpConfig(
        host=smtp_host,
        port=smtp_port,
        user=smtp_user,
        password=smtp_pass,
        from_addr=smtp_from,
    )

    stats = send_client_nav_emails(
        trade_date=effective_trade_date,
        summary_xlsx=summary_xlsx,
        smtp_config=smtp,
    )
    print(f"[OK] Client notifications: {stats['sent']} sent, {stats['skipped']} skipped, {stats.get('total', 0)} total")
    return 0


def _cmd_reconcile(args: argparse.Namespace) -> int:
    """交叉核对：持仓数量 vs 成交流水净额。

    默认 delta 模式（相邻估值日持仓变动 vs 区间成交净额），
    可选 cumulative 模式（每日持仓 vs 累计净额）。
    """
    load_env()
    from fundadmin.clients.config import NAME_TO_PRODCODE
    from fundadmin.portfolio.reconcile import (
        reconcile_around_date,
        reconcile_holding_deltas,
        reconcile_positions_vs_trades,
    )

    raw = str(getattr(args, "product", "") or "").strip()
    product_code = NAME_TO_PRODCODE.get(raw, raw)  # 接受产品名或代码
    if not product_code:
        print("[ERROR] --product 必填（产品名如 铂金8号，或代码如 SXQ602）")
        return 1

    mode = str(getattr(args, "mode", "delta") or "delta")
    center = str(getattr(args, "around", "") or "").strip()
    window = int(getattr(args, "window", 2) or 2)
    only_issues = bool(getattr(args, "only_issues", False))

    if center:
        df = reconcile_around_date(product_code, center, window=window, mode=mode)
    elif mode == "delta":
        df = reconcile_holding_deltas(product_code)
    else:
        df = reconcile_positions_vs_trades(product_code)

    if df.empty:
        print(f"[OK] {product_code}: 无可核对数据（缺少持仓或成交）")
        return 0

    if only_issues:
        df = df[df["status"] != "ok"]
        if df.empty:
            print(f"[OK] {product_code}: 全部一致，无差异")
            return 0

    n_mismatch = int((df["status"] != "ok").sum())
    print(f"[{'WARN' if n_mismatch else 'OK'}] {product_code} reconcile ({mode}): "
          f"{len(df)} row(s), {n_mismatch} issue(s)")
    with pd.option_context("display.max_rows", None, "display.width", 200):
        print(df.to_string(index=False))

    out = str(getattr(args, "out_csv", "") or "").strip()
    if out:
        Path(out).parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(out, index=False, encoding="utf-8-sig")
        print(f"[OK] written: {out}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser("fund_portfolio")
    sub = parser.add_subparsers(dest="cmd", required=True)

    parser_email_sync = sub.add_parser("email-sync", help="Pull Excel attachments from mailbox via IMAP.")
    parser_email_sync.add_argument("--trade-date", required=True, help="YYYY-MM-DD or latest")
    parser_email_sync.add_argument("--out-dir", default="", help="Output dir for downloaded attachments.")
    parser_email_sync.add_argument("--imap-host", default="", help="Override IMAP_HOST.")
    parser_email_sync.add_argument("--imap-port", default=0, type=int, help="Override IMAP_PORT.")
    parser_email_sync.add_argument("--imap-user", default="", help="Override IMAP_USER.")
    parser_email_sync.add_argument("--imap-pass", default="", help="Override IMAP_PASS.")
    parser_email_sync.add_argument("--imap-mailbox", default="", help="Override IMAP_MAILBOX.")
    parser_email_sync.add_argument("--imap-ssl", action=argparse.BooleanOptionalAction, default=True)
    parser_email_sync.add_argument("--sender-allowlist", default="", help="Comma-separated sender keywords to allow.")
    parser_email_sync.add_argument("--subject-keywords", default="", help="Comma-separated subject keywords to match.")
    parser_email_sync.add_argument("--product-scope", action=argparse.BooleanOptionalAction, default=True, help="Only sync configured fund-product emails and attachments.")
    parser_email_sync.add_argument("--scope-from-products", action=argparse.BooleanOptionalAction, default=True, help="Push PRODUCT_CONFIG codes into IMAP server-side SUBJECT SEARCH to reduce traffic. Default on.")
    parser_email_sync.add_argument("--state-file", default="", help="Processed-mail state file path. Use 'default' for the repo default path.")
    parser_email_sync.add_argument("--skip-processed", action=argparse.BooleanOptionalAction, default=False, help="Skip messages already recorded in --state-file.")
    parser_email_sync.add_argument("--print-search", action="store_true", default=False, help="Build and print the IMAP SEARCH expression for the given trade-date, then exit without contacting the server.")
    parser_email_sync.set_defaults(func=_cmd_email_sync)

    parser_sync = sub.add_parser("sync-latest", help="Incrementally sync recent trade dates and optionally build reports.")
    parser_sync.add_argument("--asof", default="", help="Anchor calendar date YYYY-MM-DD. Defaults to today.")
    parser_sync.add_argument("--lookback", type=int, default=DEFAULT_SYNC_LOOKBACK, help="Number of latest trade dates to sync.")
    parser_sync.add_argument("--inbox-root", default="", help="Root directory for date-partitioned downloaded attachments.")
    parser_sync.add_argument("--report-root", default="", help="Root directory for generated portfolio reports.")
    parser_sync.add_argument("--state-file", default="default", help="Processed-mail state file path. Defaults to the repo inbox state file.")
    parser_sync.add_argument("--build", action=argparse.BooleanOptionalAction, default=True, help="Build per-product reports and summary after syncing each trade date.")
    parser_sync.add_argument("--imap-host", default="", help="Override IMAP_HOST.")
    parser_sync.add_argument("--imap-port", default=0, type=int, help="Override IMAP_PORT.")
    parser_sync.add_argument("--imap-user", default="", help="Override IMAP_USER.")
    parser_sync.add_argument("--imap-pass", default="", help="Override IMAP_PASS.")
    parser_sync.add_argument("--imap-mailbox", default="", help="Override IMAP_MAILBOX.")
    parser_sync.add_argument("--imap-ssl", action=argparse.BooleanOptionalAction, default=True)
    parser_sync.add_argument("--sender-allowlist", default="", help="Comma-separated sender keywords to allow.")
    parser_sync.add_argument("--subject-keywords", default="", help="Comma-separated subject keywords to match.")
    parser_sync.add_argument("--product-scope", action=argparse.BooleanOptionalAction, default=True, help="Only sync configured fund-product emails and attachments.")
    parser_sync.add_argument("--scope-from-products", action=argparse.BooleanOptionalAction, default=True, help="Push PRODUCT_CONFIG codes into IMAP server-side SUBJECT SEARCH to reduce traffic. Default on.")
    parser_sync.add_argument("--skip-processed", action=argparse.BooleanOptionalAction, default=True, help="Skip messages already recorded in the processed-mail state file.")
    parser_sync.add_argument("--with-email", action="store_true", default=False, help="Send internal matrix summary email after each successful, data-complete build.")
    parser_sync.add_argument("--notify-clients", action="store_true", default=False, help="Send client NAV notification emails after each successful, data-complete build.")
    parser_sync.add_argument("--email-to", default="", help="Comma-separated matrix-email recipients (or set EMAIL_TO env var).")
    parser_sync.add_argument("--republish", action="store_true", default=False, help="Ignore published-state and resend matrix/client emails for in-window trade dates (manual re-send).")
    parser_sync.set_defaults(func=_cmd_sync_latest)

    parser_build = sub.add_parser("build", help="Compatibility alias of build-products. Prefer build-products.")
    parser_build.add_argument("--trade-date", required=True, help="YYYY-MM-DD or latest")
    parser_build.add_argument("--inbox-dir", default="", help="Downloaded attachment dir for that trade date.")
    parser_build.add_argument("--out-dir", default="", help="Output directory for product reports.")
    parser_build.add_argument("--out-xlsx", default="", help="Legacy summary alias path. Prefer --out-dir.")
    parser_build.add_argument("--with-charts", action="store_true", default=False, help="Generate holdings pie charts for each product.")
    parser_build.add_argument("--with-email", action="store_true", default=False, help="Send summary email after build.")
    parser_build.add_argument("--email-to", default="", help="Comma-separated recipient addresses (or set EMAIL_TO env var).")
    parser_build.add_argument("--smtp-host", default="", help="Override SMTP_HOST.")
    parser_build.add_argument("--smtp-port", default=0, type=int, help="Override SMTP_PORT.")
    parser_build.add_argument("--smtp-user", default="", help="Override SMTP_USER.")
    parser_build.add_argument("--smtp-pass", default="", help="Override SMTP_PASS.")
    parser_build.add_argument("--smtp-from", default="", help="Override EMAIL_FROM.")
    parser_build.set_defaults(func=_cmd_build)

    parser_xb = sub.add_parser("build-cross-broker", help="Build cross-broker portfolio report (CICC + CITIC USD/HKD + SWHYSC domestic).")
    parser_xb.add_argument("--trade-date", required=True, help="YYYY-MM-DD or latest")
    parser_xb.add_argument("--valuation-path", default="", help="Path to valuation CSV/Excel for NAV.")
    parser_xb.add_argument("--cicc-paths", default="", help="Comma-separated CICC holdings CSV/Excel paths.")
    parser_xb.add_argument("--citic-usd-underlying-paths", default="", help="Comma-separated CITIC USD Underlying paths.")
    parser_xb.add_argument("--citic-usd-balance-paths", default="", help="Comma-separated CITIC USD Balance paths (for FX).")
    parser_xb.add_argument("--citic-hkd-underlying-paths", default="", help="Comma-separated CITIC HKD Underlying paths.")
    parser_xb.add_argument("--citic-hkd-balance-paths", default="", help="Comma-separated CITIC HKD Balance paths (for FX).")
    parser_xb.add_argument("--swhysc-valuation-paths", default="", help="Comma-separated Shenwan SWHYSC valuation Excel paths.")
    parser_xb.add_argument("--out-xlsx", default="", help="Output xlsx path.")
    parser_xb.set_defaults(func=_cmd_build_cross_broker)

    parser_bp = sub.add_parser("build-products", help="Build portfolio reports per product (auto-group files from inbox).")
    parser_bp.add_argument("--trade-date", required=True, help="YYYY-MM-DD or latest")
    parser_bp.add_argument("--inbox-dir", default="", help="Downloaded attachment dir for that trade date.")
    parser_bp.add_argument("--out-dir", default="", help="Output directory for product reports.")
    parser_bp.add_argument("--with-charts", action="store_true", default=False, help="Generate holdings pie charts for each product.")
    parser_bp.add_argument("--with-email", action="store_true", default=False, help="Send summary email after build.")
    parser_bp.add_argument("--notify-clients", action="store_true", default=False, help="Send client NAV notification emails after a successful, data-complete build.")
    parser_bp.add_argument("--email-to", default="", help="Comma-separated recipient addresses (or set EMAIL_TO env var).")
    parser_bp.add_argument("--smtp-host", default="", help="Override SMTP_HOST.")
    parser_bp.add_argument("--smtp-port", default=0, type=int, help="Override SMTP_PORT.")
    parser_bp.add_argument("--smtp-user", default="", help="Override SMTP_USER.")
    parser_bp.add_argument("--smtp-pass", default="", help="Override SMTP_PASS.")
    parser_bp.add_argument("--smtp-from", default="", help="Override EMAIL_FROM.")
    parser_bp.set_defaults(func=_cmd_build_products)

    parser_prune = sub.add_parser(
        "prune-inbox",
        help="Delete old fund_inbox/<date>/ subdirs beyond a keep-last window. Manual only; never auto-runs.",
    )
    parser_prune.add_argument("--inbox-root", default="", help="Root of fund_inbox directory. Defaults to outputs/excels/fund_inbox/.")
    parser_prune.add_argument("--keep-last", type=int, default=30, help="Keep this many most-recent date dirs. Default 30.")
    parser_prune.add_argument("--dry-run", action="store_true", default=False, help="Only print what would be deleted, no actual delete.")
    parser_prune.set_defaults(func=_cmd_prune_inbox)

    parser_nc = sub.add_parser("notify-clients", help="Send NAV notification emails to clients per product.")
    parser_nc.add_argument("--trade-date", required=True, help="YYYY-MM-DD or latest")
    parser_nc.add_argument("--summary-xlsx", default="", help="Path to fund_portfolio_summary Excel. Defaults to outputs/reports/fund_portfolios/<date>/fund_portfolio_summary_<date>.xlsx")
    parser_nc.add_argument("--smtp-host", default="", help="Override SMTP_HOST.")
    parser_nc.add_argument("--smtp-port", default=0, type=int, help="Override SMTP_PORT.")
    parser_nc.add_argument("--smtp-user", default="", help="Override SMTP_USER.")
    parser_nc.add_argument("--smtp-pass", default="", help="Override SMTP_PASS.")
    parser_nc.add_argument("--smtp-from", default="", help="Override EMAIL_FROM.")
    parser_nc.set_defaults(func=_cmd_notify_clients)

    parser_rec = sub.add_parser(
        "reconcile",
        help="Cross-check holdings quantity vs transaction net (持仓 vs 成交流水交叉核对).",
    )
    parser_rec.add_argument("--product", required=True, help="产品名（如 铂金8号）或代码（如 SXQ602）。")
    parser_rec.add_argument("--mode", choices=["delta", "cumulative"], default="delta",
                            help="delta=相邻日持仓变动 vs 区间成交净额（默认）；cumulative=每日持仓 vs 累计净额。")
    parser_rec.add_argument("--around", default="", help="以该估值日为中心核对前后 ±window 个交易日（YYYY-MM-DD）。")
    parser_rec.add_argument("--window", type=int, default=2, help="--around 的前后窗口大小，默认 2。")
    parser_rec.add_argument("--only-issues", action="store_true", default=False, help="仅显示 status != ok 的差异行。")
    parser_rec.add_argument("--out-csv", default="", help="可选：将核对明细写出为 CSV。")
    parser_rec.set_defaults(func=_cmd_reconcile)

    args = parser.parse_args(argv)
    func = getattr(args, "func", None)
    if not callable(func):
        raise RuntimeError("command handler missing")
    return int(func(args))


if __name__ == "__main__":
    raise SystemExit(main())
