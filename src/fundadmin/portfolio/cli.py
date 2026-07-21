"""CLI 命令注册与 argv 分发（薄壳）。从 operations.py 拆分（god module 治理）。

职责:
  - 定义 argparse 子命令与参数
  - ``_cmd_*`` 薄函数：解析 args → 调 operations 函数 → return int
  - 不包含 IMAP、解析、持久化、报表渲染等业务逻辑
"""

from __future__ import annotations

import argparse
import logging
import warnings
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

from fundadmin.core.config import get_env, load_env
from fundadmin.notifications.email import SmtpConfig
from fundadmin.portfolio.cross_broker_report import CrossBrokerInput, build_cross_broker_report
from fundadmin.portfolio.email_filters import build_imap_search_criteria
from fundadmin.portfolio.maintenance import format_bytes, prune_inbox
from fundadmin.portfolio.client_notifier import send_client_nav_emails
from fundadmin.portfolio.operations import (
    DEFAULT_SYNC_LOOKBACK,
    PUBLISH_SCAN_DAYS,
    _build_product_reports_for_trade_date,
    _default_inbox_root,
    _default_mail_state_path,
    _default_publish_state_path,
    _default_report_root,
    _load_publish_state,
    _mark_publish_channel,
    _parse_csv_arg,
    _parse_optional_ymd,
    _publish_channel_done,
    _resolve_latest_trade_dates,
    _resolve_sender_tokens,
    _resolve_state_file_arg,
    _resolve_summary_xlsx_for_trade_date,
    _run_email_sync_for_trade_date,
    _save_publish_state,
)

logger = logging.getLogger(__name__)


def _resolve_trade_date_arg(value: str, *, today: date | None = None) -> date:
    raw = str(value or "").strip().lower()
    if not raw:
        raise ValueError("--trade-date is required")
    if raw == "latest":
        resolved = (today or date.today()) - timedelta(days=1)
        return resolved
    from fundadmin.portfolio.operations import _parse_ymd

    return _parse_ymd(raw)


def _split_paths(value: str) -> list[Path]:
    return [Path(p.strip()) for p in str(value or "").split(",") if p.strip()]


def _warn_cli_secret(args: argparse.Namespace) -> None:
    for attr, env_name in (("imap_pass", "IMAP_PASS"), ("smtp_pass", "SMTP_PASS")):
        if str(getattr(args, attr, "") or "").strip():
            flag = "--" + attr.replace("_", "-")
            warnings.warn(
                f"通过 {flag} 传入明文口令不安全（会出现在进程表与 shell 历史中）；"
                f"请改为在 .env 设置 {env_name}，并去掉该命令行参数（security-4）。",
                stacklevel=2,
            )


def _cmd_email_sync(args: argparse.Namespace) -> int:
    load_env()
    trade_date = _resolve_trade_date_arg(args.trade_date)
    if bool(getattr(args, "print_search", False)):
        sender_tokens = _resolve_sender_tokens(args)
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
    state_file = _resolve_state_file_arg(state_value, inbox_root=_default_inbox_root())
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
    from fundadmin.portfolio.operations import _extract_trade_dates_from_path, _parse_ymd

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
    import shutil
    from fundadmin.portfolio.operations import _ensure_dir

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

    total_saved = 0
    total_skipped = 0
    download_failures = 0
    for received_date in resolved_dates:
        out_dir = inbox_root / received_date.isoformat()
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
                if not want_matrix and not want_clients:
                    continue
            else:
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
            if want_clients:
                c_sent = int(payload.get("client_notify_sent") or 0)
                c_failed = int(payload.get("client_notify_failed") or 0)
                skip_reason = str(payload.get("client_notify_skip_reason") or "")
                if skip_reason:
                    pass
                elif c_failed == 0:
                    _mark_publish_channel(publish_state, trade_date, "clients")
                    clients_sent += 1
                elif c_sent == 0:
                    logger.error("client NAV notify 全部失败 %s，将于下次运行重试", trade_date.isoformat())
                else:
                    logger.error(
                        "client NAV notify 部分失败 %s (sent=%d failed=%d)，已标记发布避免重复，需人工补发失败客户",
                        trade_date.isoformat(), c_sent, c_failed,
                    )
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
    if download_failures and download_failures == len(resolved_dates):
        return 1
    return 0


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
    load_env()
    from fundadmin.clients.config import NAME_TO_PRODCODE
    from fundadmin.portfolio.reconcile import (
        reconcile_around_date,
        reconcile_holding_deltas,
        reconcile_positions_vs_trades,
    )

    raw = str(getattr(args, "product", "") or "").strip()
    product_code = NAME_TO_PRODCODE.get(raw, raw)
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
    parser_email_sync.add_argument("--imap-user", default="", help="Override IMAP_USER.  不安全：明文口令会泄露到进程表/历史，优先用 .env。")
    parser_email_sync.add_argument("--imap-port", default=0, type=int, help="Override IMAP_PORT.")
    parser_email_sync.add_argument("--imap-pass", default="", help="Override IMAP_PASS. 不安全：明文口令会泄露到进程表/历史，优先用 .env。")
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
    parser_sync.add_argument("--imap-pass", default="", help="Override IMAP_PASS. 不安全：明文口令会泄露到进程表/历史，优先用 .env。")
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
    parser_build.add_argument("--smtp-pass", default="", help="Override SMTP_PASS. 不安全：明文口令会泄露到进程表/历史，优先用 .env。")
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
    parser_bp.add_argument("--smtp-pass", default="", help="Override SMTP_PASS. 不安全：明文口令会泄露到进程表/历史，优先用 .env。")
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
    parser_nc.add_argument("--smtp-pass", default="", help="Override SMTP_PASS. 不安全：明文口令会泄露到进程表/历史，优先用 .env。")
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
    _warn_cli_secret(args)
    func = getattr(args, "func", None)
    if not callable(func):
        raise RuntimeError("command handler missing")
    return int(func(args))


if __name__ == "__main__":
    raise SystemExit(main())
