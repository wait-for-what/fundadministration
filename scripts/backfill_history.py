"""一次性历史回填：从邮箱往前下载券商附件，按交易日重建分券商持仓入库。

特点:
- 复用 operations 的内部函数（IMAP 增量下载 + 跨收件日按交易日分组 + 构建/入库）。
- 全程**不发任何邮件**（with_email=False, notify_clients=False）。
- 下载段按 message-id 状态去重，可反复执行；构建段按 (产品,估值日) 删旧重写，幂等，
  并把历史遗留的 broker='' 折叠行修正为分券商行。

用法:
    python scripts/backfill_history.py --start 2025-06-01 [--end 2026-06-16]
    python scripts/backfill_history.py --start 2025-06-01 --download-only
    python scripts/backfill_history.py --build-only           # 仅用已下载的 inbox 重建入库
"""

from __future__ import annotations

import argparse
import logging
import socket
from datetime import date, timedelta
from pathlib import Path

# 给所有 socket（含 imaplib）设全局超时：避免某个收件日在 IMAP fetch 上无限挂起，
# 超时即抛异常被单日 try/except 捕获，回填继续往后跑。
socket.setdefaulttimeout(180)

from fundadmin.core.config import get_env, load_env
from fundadmin.portfolio.operations import (
    ImapConfig,
    _build_product_reports_for_trade_date,
    _collect_inbox_files_by_trade_date,
    _default_inbox_root,
    _default_mail_state_path,
    _default_report_root,
    _fetch_excel_attachments_via_imap_with_state,
    _stage_trade_date_files,
)

logger = logging.getLogger("backfill")


def _daterange(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def main() -> int:
    ap = argparse.ArgumentParser(description="历史回填：邮箱附件 -> 分券商持仓入库（不发邮件）")
    ap.add_argument("--start", required=True, help="起始收件日 YYYY-MM-DD")
    ap.add_argument("--end", default=date.today().isoformat(), help="结束收件日 YYYY-MM-DD（默认今天）")
    ap.add_argument("--download-only", action="store_true", help="只下载附件，不构建入库")
    ap.add_argument("--build-only", action="store_true", help="跳过下载，仅用已下载 inbox 重建入库")
    ap.add_argument("--include-weekends", action="store_true", help="下载段也扫描周末收件日")
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

    # ---- 下载段 ----
    if not args.build_only:
        total_saved = 0
        days = 0
        for d in _daterange(start, end):
            if d.weekday() >= 5 and not args.include_weekends:
                continue
            days += 1
            out_dir = inbox_root / d.isoformat()
            try:
                res = _fetch_excel_attachments_via_imap_with_state(
                    imap=imap,
                    target_date=d,
                    out_dir=out_dir,
                    sender_allowlist=None,
                    subject_keywords=None,
                    state_file=state_file,
                    skip_processed=True,
                    product_scope=True,
                    scope_from_products=True,
                )
                if res.saved_paths or res.matched_messages:
                    total_saved += len(res.saved_paths)
                    print(
                        f"[dl] {d.isoformat()} saved={len(res.saved_paths)} "
                        f"matched={res.matched_messages} skipped={res.skipped_messages}",
                        flush=True,
                    )
            except Exception as exc:  # noqa: BLE001 - 单日失败不阻断整体回填
                print(f"[dl] {d.isoformat()} ERR {type(exc).__name__}: {exc}", flush=True)
        print(f"[dl] DONE scanned_days={days} total_new_files={total_saved}", flush=True)

    # ---- 构建/入库段 ----
    if not args.download_only:
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

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
