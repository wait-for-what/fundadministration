"""业务函数：拉 Excel、解析、汇总、推送。

治理提示::
    从 god module 拆分出 cli.py（命令注册）和 reports.py（报表/邮件完整性检查）。
    operations.py 保留 IMAP、解析、持久化、报表编排等业务逻辑，不直接处理 argparse。
"""

from __future__ import annotations

import hashlib
import imaplib
import json
import logging
import os
import re
import shutil
import ssl
import tempfile
import time
import warnings
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
from fundadmin.portfolio.cross_broker_report import (
    PRODUCT_CONFIG,
    build_product_reports,
    build_summary_excel,
    score_product_inputs_for_date,
)
from fundadmin.portfolio.reports import (
    client_unit_nav_missing,
    product_email_completion_issues,
)
from fundadmin.portfolio.email_filters import (
    build_imap_search_criteria,
    is_target_fund_attachment,
    is_target_fund_email,
)
from fundadmin.portfolio.by_broker_email import send_by_broker_summary_email
from fundadmin.portfolio.client_notifier import send_client_nav_emails
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


def _sender_allowed(sender: str, allow_tokens: set[str]) -> bool:
    """发件人是否在白名单内。

    只对解析出的邮件地址（addr-spec）做匹配，不匹配显示名——否则把可信关键词放进
    显示名即可冒充（email-ingest-1）。allow_tokens 为空时不做过滤（返回 True）。
    """
    if not allow_tokens:
        return True
    from email.utils import parseaddr

    addr = (parseaddr(sender or "")[1] or "").lower()
    if not addr:
        return False
    return any(tok in addr for tok in allow_tokens)


def _resolve_sender_tokens(args: argparse.Namespace) -> list[str]:
    """合并 CLI --sender-allowlist 与 .env 的 IMAP_SENDER_ALLOWLIST。

    让每日定时任务无需改 plist 即可在 .env 配置可信发件人域名/地址。两者皆空时告警，
    使"入库无发件人鉴别"对运维可见（email-ingest-1）。
    """
    tokens = _parse_csv_arg(str(getattr(args, "sender_allowlist", "") or ""))
    if not tokens:
        tokens = _parse_csv_arg(str(get_env("IMAP_SENDER_ALLOWLIST", default="") or ""))
    if not tokens:
        warnings.warn(
            "未配置发件人白名单（--sender-allowlist 或 .env IMAP_SENDER_ALLOWLIST 均为空）："
            "券商邮件入库无发件人鉴别，任何落入邮箱且主题命中产品名的邮件都会被采信（email-ingest-1）。",
            stacklevel=2,
        )
    return tokens


def _atomic_write_text(path: Path, text: str) -> None:
    """原子写文本：写同目录临时文件并 fsync，再 os.replace 覆盖目标。

    ``Path.write_text`` 是"截断-再写"，崩溃/掉电会留下半截 JSON，损坏去重/发布
    状态（ops-sync-3 / reliability-2）。os.replace 在同一文件系统上是原子替换。
    """
    _ensure_dir(path.parent)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=path.name + ".", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(text)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _load_email_sync_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"version": 1, "processed_messages": {}}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            raise ValueError("payload is not a dict")
    except (json.JSONDecodeError, ValueError, OSError) as exc:
        # 去重状态损坏只影响"是否重新扫描"，不触发对外发送：降级为空状态重新扫描，
        # 比让整轮下载因一处坏文件而崩溃更安全（ops-sync-3）。
        logger.error("email_sync_state 损坏，按空状态重建并重新扫描: %s (%s)", path, exc)
        return {"version": 1, "processed_messages": {}}
    processed = raw.get("processed_messages")
    if not isinstance(processed, dict):
        raw["processed_messages"] = {}
    raw["version"] = int(raw.get("version") or 1)
    return raw


def _save_email_sync_state(path: Path, state: dict[str, Any]) -> None:
    _atomic_write_text(path, json.dumps(state, ensure_ascii=False, indent=2))


def _load_publish_state(path: Path) -> dict[str, Any]:
    """发布记录：trade_date -> {"matrix_sent_at": iso, "clients_sent_at": iso}。

    用于 sync-latest 发布段去重：同一交易日某渠道已成功发送后不再重发，
    数据不齐的交易日不会写入记录，留待后续补料后自动补发。
    """
    if not path.exists():
        return {"version": 1, "published": {}}
    # 发布状态是"发送去重"凭据。与 email_sync_state 不同，损坏时绝不静默重置为空——
    # 那会导致对全体客户重发净值邮件。宁可让本轮发布段报错由运维介入（ops-sync-3）。
    raw = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError(f"invalid publish state payload: {path}")
    if not isinstance(raw.get("published"), dict):
        raw["published"] = {}
    raw["version"] = int(raw.get("version") or 1)
    return raw


def _save_publish_state(path: Path, state: dict[str, Any]) -> None:
    _atomic_write_text(path, json.dumps(state, ensure_ascii=False, indent=2))


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
                if not _sender_allowed(sender, allow_senders):
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
                if not _sender_allowed(sender, allow_senders):
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
        sender_allowlist=_resolve_sender_tokens(args),
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
    from fundadmin.clients.store import insert_attachment_with_rows
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
            def _build_rows(ingest_id: int, _sheets=sheets) -> pd.DataFrame:
                rows: list[dict[str, Any]] = []
                for s_idx, (s_name, df) in enumerate(_sheets):
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
                return pd.DataFrame(rows)

            # 登记附件 + 写原始行在同一事务内完成，避免孤儿附件（raw-ingest gap）。
            _ingest_id, is_new, _n_rows = insert_attachment_with_rows(meta, _build_rows)
            if not is_new:
                skipped += 1
                continue
            ingested += 1
        except Exception:
            logger.exception("raw-layer ingest failed for %s", path)

    print(f"[OK] raw-layer ingest: {ingested} new attachment(s), {skipped} duplicate(s) skipped")


def _persist_curated_layer(
    results: list[dict[str, Any]],
    *,
    effective_trade_date: date,
) -> bool:
    """核心结构层：写 product_valuation + fund_positions（分券商）。

    返回 True 表示入库成功；False 表示发生异常（已记日志）。调用方据此阻断通知，
    避免在结构层写失败、DB 处于不一致状态时仍对外发净值邮件（ops-positions-2）。

    fund_positions 按 (产品, 标的, 券商) 逐行落地：同一标的若同时在中金、中信持有，
    各券商单独成行（broker 真实填充），不再跨券商折叠。写入前先删除该产品该估值日的
    旧快照，保证重跑幂等、不与旧 broker='' 折叠行重复计数。
    """
    from fundadmin.clients.config import NAME_TO_PRODCODE
    from fundadmin.clients.store import (
        replace_positions,
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
            # 单事务内"清旧快照 + 写新分券商行"：重跑幂等，且失败回滚到旧快照，
            # 不留空/半快照（clients-store-3 / ops-sync-11）。
            n_pos += replace_positions(pcode, as_of, pd.concat(frames, ignore_index=True))
        print(f"[OK] curated-layer: {n_val} valuation row(s), {n_pos} position row(s, per-broker)")
        return True
    except Exception:
        logger.exception("curated-layer persistence failed")
        return False


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

    # 核心结构层：从 build 结果写 product_valuation + fund_positions。
    # 入库失败时 curated_ok=False，后续据此阻断对外通知（ops-positions-2）。
    curated_ok = _persist_curated_layer(results, effective_trade_date=effective_trade_date)

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
    client_notify_eligible = 0
    client_notify_failed = 0
    client_notify_skip_reason = ""
    # 矩阵邮件与客户净值邮件分别独立把门：内部矩阵用 asset_nav 即可，客户邮件还需
    # unit_nav（notifications-2）；且结构层入库失败时两者都不发（ops-positions-2）。
    if with_email or notify_clients:
        completion_issues = product_email_completion_issues(
            results=results,
            summary_path=summary_path,
            require_charts=with_charts,
            chart_paths=chart_paths,
        )
        if not curated_ok:
            completion_issues.append("curated-layer persistence failed (DB write incomplete)")

        client_issues = list(completion_issues)
        missing_unit_nav = client_unit_nav_missing(results)
        if missing_unit_nav:
            client_issues.append(
                f"missing unit_nav (client market value uncomputable): {', '.join(missing_unit_nav)}"
            )

        matrix_blocked = bool(completion_issues)
        client_blocked = bool(client_issues)
        if with_email and matrix_blocked:
            email_skip_reason = "; ".join(completion_issues)
            print(f"[WARN] matrix email skipped: {email_skip_reason}")
        if notify_clients and client_blocked:
            client_notify_skip_reason = "; ".join(client_issues)
            print(f"[WARN] client NAV notify skipped: {client_notify_skip_reason}")

        send_matrix = with_email and not matrix_blocked
        send_clients = notify_clients and not client_blocked
        if send_matrix or send_clients:
            load_env()

            # ---- 内部持仓汇总邮件（EMAIL_TO，QQ 邮箱 SMTP）----
            if send_matrix:
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
            if send_clients:
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
                    client_notify_eligible = int(stats.get("eligible", 0))
                    client_notify_failed = int(stats.get("failed", 0))
                    print(
                        f"[OK] client NAV notify: {stats['sent']} sent, "
                        f"{stats['skipped']} skipped, {stats.get('failed', 0)} failed, "
                        f"{stats.get('total', 0)} total"
                    )
                    if client_notify_failed:
                        # 发送失败的客户保留以便重试（notifications-4）：发布状态由调用方据此决定。
                        print(
                            f"[WARN] client NAV notify: {client_notify_failed} 个客户发送失败 "
                            f"-> {', '.join(stats.get('failed_clients', []))}"
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
        payload["client_notify_eligible"] = client_notify_eligible
        payload["client_notify_failed"] = client_notify_failed
        payload["client_notify_skip_reason"] = client_notify_skip_reason
    return payload






