"""全量归档下载器：抗 QQ 掉线（断点续传 + 自动重连 + 分块），只下载不入库。

背景:
- QQ 的 IMAP SEARCH 忽略所有条件(日期/主题)，永远返回整箱；且长时间批量抓取会被
  QQ 主动掐断(socket EOF)。因此必须客户端遍历全箱头、is_target_fund_email 过滤，
  逐封抓正文存附件，并对每个 message-id 记 state，崩了能续传。

行为:
- 遍历 INBOX 全部邮件(按 seq 分块取头)，命中券商报告的逐封抓 BODY[] 提附件，
  按收件日(INTERNALDATE)存到 <out-dir>/<yyyy-mm-dd>/。
- 处理过的 message-id 落 state JSON；重跑自动跳过；任一 IMAP 异常重连退避重试。
- 只下载。入库另跑:
    python scripts/backfill_history.py --build-only --start 2020-08-01 --inbox-root <out-dir>
"""

from __future__ import annotations

import argparse
import hashlib
import imaplib
import json
import re
import socket
import ssl
import time
from datetime import date
from email import message_from_bytes
from pathlib import Path

from fundadmin.core.config import get_env, load_env
from fundadmin.portfolio.email_filters import is_target_fund_attachment, is_target_fund_email
from fundadmin.portfolio.operations import (
    SUPPORTED_ATTACHMENT_SUFFIXES,
    ImapConfig,
    _decode_email_header,
    _iter_attachments,
    _safe_filename,
)

socket.setdefaulttimeout(120)
_MON = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
        "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}


def _imap() -> ImapConfig:
    return ImapConfig(
        host=str(get_env("IMAP_HOST", required=True) or ""),
        user=str(get_env("IMAP_USER", required=True) or ""),
        password=str(get_env("IMAP_PASS", required=True) or ""),
        mailbox=str(get_env("IMAP_MAILBOX", "INBOX") or "INBOX"),
        port=int(get_env("IMAP_PORT", "993") or "993"),
        use_ssl=True,
    )


def _connect(imap: ImapConfig) -> imaplib.IMAP4:
    cli = imaplib.IMAP4_SSL(imap.host, imap.port, ssl_context=ssl.create_default_context())
    cli.login(imap.user, imap.password)
    st, _ = cli.select(imap.mailbox, readonly=True)
    if st != "OK":
        cli.select("收件箱", readonly=True)
    return cli


_IMAP_ERRS = (imaplib.IMAP4.abort, imaplib.IMAP4.error, OSError, ssl.SSLError)


def _fetch_retry(state: dict, imap: ImapConfig, seq: str, parts: str):
    """带重连退避的 FETCH。state['cli'] 持有当前连接，断了就重连。返回 data。"""
    for attempt in range(8):
        try:
            typ, data = state["cli"].fetch(seq, parts)
            if typ != "OK":
                raise imaplib.IMAP4.abort(f"FETCH status {typ}")
            return data
        except _IMAP_ERRS as exc:
            wait = min(60, 2 ** attempt)
            print(f"[archive] FETCH 失败(seq={seq[:24]}.. attempt={attempt}): {type(exc).__name__}: {exc}; {wait}s 后重连", flush=True)
            try:
                state["cli"].logout()
            except Exception:
                pass
            time.sleep(wait)
            try:
                state["cli"] = _connect(imap)
            except Exception as exc2:
                print(f"[archive] 重连失败: {exc2}", flush=True)
    return None


def _internaldate(meta: str) -> date | None:
    m = re.search(r'INTERNALDATE "(\d{1,2})-(\w{3})-(\d{4})', meta)
    if not m:
        return None
    return date(int(m.group(3)), _MON.get(m.group(2), 1), int(m.group(1)))


def main() -> int:
    ap = argparse.ArgumentParser(description="全量归档下载器(抗掉线/续传)")
    ap.add_argument("--out-dir", default="outputs/excels/backfill_inbox")
    ap.add_argument("--state", default="outputs/excels/backfill_inbox/_state/archive_state.json")
    ap.add_argument("--chunk", type=int, default=100, help="每次取头的邮件数")
    ap.add_argument("--since-year", type=int, default=0, help="只处理收件年份 >= 该值")
    ap.add_argument("--max", type=int, default=0, help="最多处理多少封(0=全部，调试用)")
    args = ap.parse_args()

    load_env()
    imap = _imap()
    out_root = Path(args.out_dir)
    out_root.mkdir(parents=True, exist_ok=True)
    state_path = Path(args.state)
    state_path.parent.mkdir(parents=True, exist_ok=True)

    processed: set[str] = set()
    if state_path.exists():
        try:
            processed = set(json.loads(state_path.read_text(encoding="utf-8")).get("processed_message_ids", []))
        except Exception:
            processed = set()
    print(f"[archive] 续传: 已处理 {len(processed)} 封", flush=True)

    st = {"cli": _connect(imap)}
    typ, d = st["cli"].search(None, "ALL")
    all_ids = d[0].split() if (typ == "OK" and d and d[0]) else []
    total = len(all_ids)
    print(f"[archive] INBOX 共 {total} 封; chunk={args.chunk} since_year={args.since_year or '全部'}", flush=True)

    def save_state() -> None:
        tmp = state_path.with_suffix(".tmp")
        tmp.write_text(json.dumps({"processed_message_ids": sorted(processed)}, ensure_ascii=False), encoding="utf-8")
        tmp.replace(state_path)

    saved_total = 0
    scanned = 0
    matched_new = 0
    i = 0
    while i < total:
        if args.max and scanned >= args.max:
            break
        chunk = all_ids[i:i + args.chunk]
        i += args.chunk
        seq = b",".join(chunk).decode()
        data = _fetch_retry(st, imap, seq, "(INTERNALDATE BODY.PEEK[HEADER])")
        if data is None:
            print(f"[archive] 跳过块 @i={i}", flush=True)
            continue

        todo = []  # (seqid, key, idate, subject)
        for it in data:
            if not (isinstance(it, tuple) and len(it) >= 2):
                continue
            meta = it[0].decode(errors="replace")
            hdr = it[1]
            sm = re.match(r"^(\d+)", meta)
            seqid = sm.group(1) if sm else None
            idate = _internaldate(meta)
            scanned += 1
            if args.since_year and idate and idate.year < args.since_year:
                continue
            msg = message_from_bytes(hdr)
            subj = _decode_email_header(msg.get("Subject"))
            frm = _decode_email_header(msg.get("From"))
            if not is_target_fund_email(subject=subj, sender=frm):
                continue
            mid = str(msg.get("Message-ID") or msg.get("Message-Id") or "").strip().lower()
            key = mid or ("hdr:" + hashlib.sha1(hdr).hexdigest())
            if key in processed:
                continue
            if seqid:
                todo.append((seqid, key, idate, subj))

        for seqid, key, idate, subj in todo:
            body = _fetch_retry(st, imap, seqid, "(BODY.PEEK[])")
            payload = None
            for it in body or []:
                if isinstance(it, tuple) and len(it) >= 2:
                    payload = it[1]
                    break
            if not payload:
                processed.add(key)  # 取不到正文也记下，避免卡死重试
                continue
            msg = message_from_bytes(payload)
            subj2 = _decode_email_header(msg.get("Subject")) or subj
            day = idate.isoformat() if idate else "undated"
            od = out_root / day
            od.mkdir(parents=True, exist_ok=True)
            n = 0
            for fname, fbytes in _iter_attachments(msg):
                if not str(fname).lower().endswith(SUPPORTED_ATTACHMENT_SUFFIXES):
                    continue
                if not is_target_fund_attachment(filename=fname, subject=subj2):
                    continue
                token = hashlib.sha1(key.encode()).hexdigest()[:12]
                base = f"{_safe_filename(subj2) or 'email'}_{token}_{n:02d}_{_safe_filename(fname)}"
                try:
                    (od / base).write_bytes(fbytes)
                    n += 1
                    saved_total += 1
                except Exception as exc:
                    print(f"[archive] 写盘失败 {base}: {exc}", flush=True)
            processed.add(key)
            matched_new += 1

        if (i // args.chunk) % 10 == 0:
            save_state()
            print(f"[archive] 进度 {min(i, total)}/{total}  本次新增命中={matched_new} 附件={saved_total} state={len(processed)}", flush=True)

    save_state()
    try:
        st["cli"].logout()
    except Exception:
        pass
    print(f"[archive] DONE scanned={scanned} new_matched={matched_new} saved_attachments={saved_total} state_total={len(processed)}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
