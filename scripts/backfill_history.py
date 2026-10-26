"""历史回填：整箱单遍扫描邮箱，按收件日下载券商附件；再用 inbox 按交易日重建入库。

设计（full-access 邮箱适配，4 万+ 封）:
- 逐日 SEARCH 会每次 peek 整箱头导致不可用（QQ SEARCH 忽略 SINCE/BEFORE 返回整箱），
  因此本脚本对**整箱只做一遍** HEADER peek，客户端按 INTERNALDATE 筛 [start, end] 收件窗口，
  按 message-key 状态去重后，把匹配邮件的附件下载到 inbox/<收件日>/（与每日同步同布局）。
- 全程**不发任何邮件**；下载与构建均可反复执行（状态去重 + 删旧重写，幂等）。
- 匹配语义与每日同步完全一致：复用 operations/email_filters 的同一批内部函数
  （发件人白名单 / 产品邮件识别 / 附件识别 / message-key 状态）。

用法:
    python scripts/backfill_history.py --start 2026-01-01 [--end 2026-08-28]
    python scripts/backfill_history.py --start 2026-01-01 --download-only
    python scripts/backfill_history.py --build-only           # 仅用已下载的 inbox 重建入库
"""

from __future__ import annotations

import argparse
import imaplib
import logging
import re
import socket
import ssl
import time
from collections.abc import Callable
from datetime import date, timedelta
from email import message_from_bytes
from email.message import Message
from pathlib import Path

# 给所有 socket（含 imaplib）设全局超时：避免单条 IMAP 命令无限挂起。
socket.setdefaulttimeout(180)

from fundadmin.core.config import get_env, load_env
from fundadmin.portfolio.cli import (
    _collect_inbox_files_by_trade_date,
    _stage_trade_date_files,
)
from fundadmin.portfolio.email_filters import (
    is_target_fund_attachment,
    is_target_fund_email,
)
from fundadmin.portfolio.operations import (
    SUPPORTED_ATTACHMENT_SUFFIXES,
    ImapConfig,
    _build_product_reports_for_trade_date,
    _build_state_entry,
    _decode_email_header,
    _default_inbox_root,
    _default_mail_state_path,
    _default_report_root,
    _ensure_dir,
    _iter_attachments,
    _load_email_sync_state,
    _message_key_from_header,
    _message_token,
    _parse_imap_internaldate,
    _safe_filename,
    _save_email_sync_state,
    _sender_allowed,
)

logger = logging.getLogger("backfill")

_HEADER_BATCH = 50
_RFC822_BATCH = 20


def _daterange(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def _connect(imap: ImapConfig) -> imaplib.IMAP4:
    """建立 IMAP 连接、登录并选中目标邮箱（含中文收件箱 fallback）。"""
    context = ssl.create_default_context()
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


def _batch_fetch(
    box: list[imaplib.IMAP4],
    mids: list[bytes],
    spec: str,
    reconnect: Callable[[], imaplib.IMAP4],
) -> dict[bytes, object]:
    """批量 FETCH（带重连与二分降批），返回 mid -> payload/meta。"""
    if not mids:
        return {}
    seq = b",".join(mids).decode("ascii")

    def _once() -> list:
        data = box[0].fetch(seq, spec)[1]
        return data or []

    try:
        data = _once()
    except (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError) as exc:
        logger.warning("IMAP fetch aborted (%d ids): %s; reconnecting", len(mids), exc)
        if len(mids) <= 1:
            return {}
        try:
            box[0].logout()
        except Exception:
            pass
        box[0] = reconnect()
        mid = len(mids) // 2
        return {
            **_batch_fetch(box, mids[:mid], spec, reconnect),
            **_batch_fetch(box, mids[mid:], spec, reconnect),
        }

    out: dict[bytes, object] = {}
    for item in data:
        if isinstance(item, tuple) and len(item) >= 2:
            meta, payload = item[0], item[1]
            if isinstance(meta, (bytes, bytearray)) and isinstance(payload, (bytes, bytearray)):
                m = re.match(rb"^(\d+)", bytes(meta))
                if m:
                    out[m.group(1)] = (bytes(meta), bytes(payload))
        elif isinstance(item, (bytes, bytearray)):
            # spec 只含元数据时（无 BODY 段）服务端可能返回裸行
            m = re.match(rb"^(\d+)\s+\(", bytes(item))
            if m:
                out[m.group(1)] = (bytes(item), b"")
    return out


def _download_pass(imap: ImapConfig, *, start: date, end: date, inbox_root: Path, state_file: Path) -> None:
    """整箱单遍：peek 全部头 → 收件窗口+白名单+产品过滤 → 按收件日下载附件。"""
    box = [_connect(imap)]
    status, data = box[0].search(None, "ALL")
    if status != "OK":
        raise RuntimeError("IMAP SEARCH ALL failed")
    ids = (data[0] or b"").split()
    print(f"[scan] mailbox={imap.mailbox} messages={len(ids)} window=[{start}, {end}]", flush=True)

    state = _load_email_sync_state(state_file)
    processed = dict(state.get("processed_messages") or {})
    allow_tokens = {
        str(t or "").strip().lower()
        for t in str(get_env("IMAP_SENDER_ALLOWLIST", default="") or "").split(",")
        if str(t or "").strip()
    }

    # ---- 阶段 1：peek 全部头，筛出窗口内的产品邮件 ----
    matched: dict[date, list[tuple[bytes, str, str]]] = {}  # received -> [(mid, message_key, subject)]
    skipped_known = 0
    for i in range(0, len(ids), _HEADER_BATCH):
        batch = ids[i : i + _HEADER_BATCH]
        metas = _batch_fetch(box, batch, "(INTERNALDATE BODY.PEEK[HEADER])", lambda: _connect(imap))
        for meta, payload in metas.values():
            received = _parse_imap_internaldate(meta)
            if received is None or received < start or received > end:
                continue
            msg = message_from_bytes(payload)
            subject = _decode_email_header(msg.get("Subject"))
            sender = _decode_email_header(msg.get("From"))
            if not _sender_allowed(sender, allow_tokens):
                continue
            if not is_target_fund_email(subject=subject, sender=sender):
                continue
            message_key = _message_key_from_header(msg, payload)
            if message_key in processed:
                skipped_known += 1
                continue
            matched.setdefault(received, []).append((meta, message_key, subject))
        if (i // _HEADER_BATCH) % 40 == 0:
            print(
                f"[scan] {i + len(batch)}/{len(ids)} headers peeked, "
                f"matched={sum(len(v) for v in matched.values())}, known={skipped_known}",
                flush=True,
            )
    total_matched = sum(len(v) for v in matched.values())
    print(f"[scan] DONE matched={total_matched} known={skipped_known} days={len(matched)}", flush=True)

    # ---- 阶段 2：按收件日批量下载整封并保存附件 ----
    saved_total = 0
    processed_count = 0
    for d in sorted(matched):
        out_dir = inbox_root / d.isoformat()
        _ensure_dir(out_dir)
        entries = matched[d]
        saved_today = 0
        for i in range(0, len(entries), _RFC822_BATCH):
            batch = entries[i : i + _RFC822_BATCH]
            full_msgs = _batch_fetch(box, [m for m, _, _ in batch], "(RFC822)", lambda: _connect(imap))
            for meta, message_key, subject_hint in batch:
                fetched = full_msgs.get(re.match(rb"^(\d+)", meta).group(1))  # type: ignore[union-attr]
                if not fetched:
                    continue
                raw_bytes = fetched[1]
                if not isinstance(raw_bytes, (bytes, bytearray)) or not raw_bytes:
                    continue
                msg: Message = message_from_bytes(raw_bytes)
                subject = _decode_email_header(msg.get("Subject")) or subject_hint
                token = _message_token(message_key)
                saved_for_message: list[Path] = []
                for index, (filename, file_payload) in enumerate(_iter_attachments(msg), start=1):
                    if not str(filename).lower().endswith(SUPPORTED_ATTACHMENT_SUFFIXES):
                        continue
                    if not is_target_fund_attachment(filename=filename, subject=subject):
                        continue
                    prefix = _safe_filename(subject) or "email"
                    base = f"{prefix}_{token}_{index:02d}_{_safe_filename(filename)}"
                    path = out_dir / base
                    path.write_bytes(file_payload)
                    saved_for_message.append(path)
                saved_today += len(saved_for_message)
                processed[message_key] = _build_state_entry(
                    message_key=message_key,
                    msg=msg,
                    trade_date=d,
                    saved_paths=saved_for_message,
                )
                processed_count += 1
        state["processed_messages"] = processed
        _save_email_sync_state(state_file, state)
        saved_total += saved_today
        print(f"[dl] {d.isoformat()} messages={len(entries)} saved={saved_today}", flush=True)
    print(f"[dl] DONE days={len(matched)} messages={processed_count} new_files={saved_total}", flush=True)


def _build_pass(*, start: date, inbox_root: Path, report_root: Path) -> None:
    """用已下载 inbox 按交易日重建入库（删旧重写，幂等）。"""
    scan_days = (date.today() - start).days + 5
    grouped = _collect_inbox_files_by_trade_date(
        inbox_root=inbox_root, asof_date=date.today(), scan_days=scan_days
    )
    built = 0
    for td in sorted(grouped):
        staged = _stage_trade_date_files(inbox_root=inbox_root, trade_date=td, files=grouped[td])
        try:
            payload = _build_product_reports_for_trade_date(
                trade_date=td,
                inbox_dir=staged,
                report_root=report_root,
                with_email=False,
                notify_clients=False,
            )
            built += 1
            print(
                f"[build] {td.isoformat()} products={payload.get('product_count')} "
                f"out={payload.get('out_dir')}",
                flush=True,
            )
        except Exception as exc:  # noqa: BLE001 - 单日失败不阻断整体回填
            print(f"[build] {td.isoformat()} ERR {type(exc).__name__}: {exc}", flush=True)
    print(f"[build] DONE trade_dates_built={built}", flush=True)


def main() -> int:
    ap = argparse.ArgumentParser(description="历史回填：邮箱附件 -> 本地入库（不发邮件）")
    ap.add_argument("--start", required=True, help="起始收件日 YYYY-MM-DD")
    ap.add_argument("--end", default=date.today().isoformat(), help="结束收件日 YYYY-MM-DD（默认今天）")
    ap.add_argument("--download-only", action="store_true", help="只下载附件，不构建入库")
    ap.add_argument("--build-only", action="store_true", help="跳过下载，仅用已下载 inbox 重建入库")
    ap.add_argument("--inbox-root", default="", help="自定义 inbox 根目录（默认 outputs/excels/fund_inbox）")
    args = ap.parse_args()

    load_env()
    logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(name)s %(message)s")

    start = date.fromisoformat(args.start)
    end = date.fromisoformat(args.end)
    inbox_root = Path(args.inbox_root) if str(args.inbox_root or "").strip() else _default_inbox_root()
    report_root = _default_report_root()
    state_file = _default_mail_state_path(inbox_root)

    imap = ImapConfig(
        host=str(get_env("IMAP_HOST", required=True) or ""),
        user=str(get_env("IMAP_USER", required=True) or ""),
        password=str(get_env("IMAP_PASS", required=True) or ""),
        mailbox=str(get_env("IMAP_MAILBOX", "INBOX") or "INBOX"),
        port=int(get_env("IMAP_PORT", "993") or "993"),
        use_ssl=True,
    )

    if not args.build_only:
        _download_pass(imap, start=start, end=end, inbox_root=inbox_root, state_file=state_file)

    if not args.download_only:
        _build_pass(start=start, inbox_root=inbox_root, report_root=report_root)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
