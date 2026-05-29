"""SMTP 邮件发送（HTML 正文 + 内联图片 + 附件）。

用途:
- 客户分产品净值通知、运维告警等 HTML 邮件发送。

输入:
- SmtpConfig（host/port/user/password/from_addr 及重试/限速参数）。

输出:
- send_html_email 成功返回 None；失败抛 EmailSendError 子类。

失败行为:
- 鉴权失败 -> SmtpAuthError；收件人被拒 -> SmtpRecipientRefusedError；
  4xx/网络抖动按 retries 重试，仍失败 -> SmtpTransientError；5xx -> SmtpPermanentError。
"""

from __future__ import annotations

import smtplib
import ssl
import threading
import time
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from random import random


class EmailSendError(RuntimeError):
    """Base error for email delivery failures."""


class SmtpAuthError(EmailSendError):
    """SMTP authentication failed."""


class SmtpRecipientRefusedError(EmailSendError):
    """All recipients were refused by the SMTP server."""


class SmtpTransientError(EmailSendError):
    """Transient SMTP/network error (may succeed after retry)."""


class SmtpPermanentError(EmailSendError):
    """Permanent SMTP error (5xx)."""


@dataclass(frozen=True)
class SmtpConfig:
    host: str
    port: int
    user: str = ""
    password: str = ""
    from_addr: str = ""
    timeout_sec: int = 30
    # Retry policy (only for transient errors)
    retries: int = 3
    backoff: float = 1.6
    jitter: float = 0.0
    # Optional throttle between send attempts (process-wide rate limit)
    min_interval_sec: float = 0.0


def _normalize_to_addrs(to_addrs: str | Iterable[str]) -> list[str]:
    if isinstance(to_addrs, str):
        parts = [x.strip() for x in to_addrs.split(",") if x.strip()]
        return parts
    return [str(x).strip() for x in to_addrs if str(x).strip()]


def _build_message(
    *,
    subject: str,
    html_body: str,
    from_addr: str,
    to_addrs: list[str],
    images: Mapping[str, bytes] | None = None,
    attachments: Mapping[str, bytes] | None = None,
) -> MIMEMultipart:
    # 最外层 mixed：支持附件 + 内联图片 + HTML 正文
    msg = MIMEMultipart("mixed")

    # 正文部分（related：HTML + 内联图片）
    if images:
        body_part = MIMEMultipart("related")
        alt = MIMEMultipart("alternative")
        body_part.attach(alt)
        alt.attach(MIMEText(html_body, "html", "utf-8"))

        for img_id, img_bytes in images.items():
            img = MIMEImage(img_bytes)
            img.add_header("Content-ID", f"<{img_id}>")
            img.add_header("Content-Disposition", "inline", filename=f"{img_id}.png")
            body_part.attach(img)
    else:
        body_part = MIMEMultipart("alternative")
        body_part.attach(MIMEText(html_body, "html", "utf-8"))

    msg.attach(body_part)

    # 附件
    if attachments:
        for filename, file_bytes in attachments.items():
            part = MIMEApplication(file_bytes)
            part.add_header("Content-Disposition", "attachment", filename=filename)
            msg.attach(part)

    msg["Subject"] = str(Header(subject, "utf-8"))
    msg["From"] = from_addr
    msg["To"] = ", ".join(to_addrs)
    return msg


_SEND_THROTTLE_LOCK = threading.Lock()
_LAST_SEND_ATTEMPT_MONO: float | None = None


def _throttle_send_attempts(min_interval_sec: float) -> None:
    global _LAST_SEND_ATTEMPT_MONO

    interval = float(min_interval_sec or 0.0)
    if interval <= 0:
        return

    with _SEND_THROTTLE_LOCK:
        now = time.monotonic()
        if _LAST_SEND_ATTEMPT_MONO is not None:
            elapsed = now - _LAST_SEND_ATTEMPT_MONO
            wait = interval - elapsed
            if wait > 0:
                time.sleep(wait)
        _LAST_SEND_ATTEMPT_MONO = time.monotonic()


def _is_transient_error(exc: BaseException) -> bool:
    if isinstance(exc, smtplib.SMTPResponseException):
        code = int(getattr(exc, "smtp_code", 0) or 0)
        # Treat 4xx as transient (server busy/temporary failure).
        return 400 <= code < 500

    if isinstance(exc, smtplib.SMTPRecipientsRefused):
        recipients = getattr(exc, "recipients", {}) or {}
        codes: list[int] = []
        for _addr, payload in recipients.items():
            if isinstance(payload, tuple) and payload:
                try:
                    codes.append(int(payload[0]))
                except Exception:
                    continue
        return bool(codes) and all(400 <= c < 500 for c in codes)

    return isinstance(
        exc,
        (
            TimeoutError,
            OSError,
            ssl.SSLError,
            smtplib.SMTPServerDisconnected,
            smtplib.SMTPConnectError,
        ),
    )


def _smtp_status_code(exc: BaseException) -> int:
    if not isinstance(exc, smtplib.SMTPResponseException):
        return 0
    try:
        return int(getattr(exc, "smtp_code", 0) or 0)
    except Exception:
        return 0


def _sleep_backoff(smtp: SmtpConfig, attempt: int) -> None:
    base = float(smtp.backoff) ** max(0, int(attempt) - 1)
    j = float(smtp.jitter or 0.0)
    if j > 0:
        base += random() * j
    time.sleep(base)


def send_html_email(
    smtp: SmtpConfig,
    *,
    subject: str,
    html_body: str,
    to_addrs: str | Iterable[str],
    images: dict[str, bytes] | None = None,
    attachments: dict[str, bytes] | None = None,
) -> None:
    """
    Send an HTML email with optional inline PNG images and file attachments.

    - Transient errors (disconnect/timeout/SSL) are retried.
    - Auth/recipient errors are raised immediately with a classified exception.
    """
    recipients = _normalize_to_addrs(to_addrs)
    if not recipients:
        raise ValueError("to_addrs is empty")

    from_addr = smtp.from_addr or smtp.user
    if not from_addr:
        raise ValueError("smtp.from_addr is empty (and smtp.user is empty)")

    msg = _build_message(
        subject=str(subject),
        html_body=str(html_body),
        from_addr=str(from_addr),
        to_addrs=recipients,
        images=images,
        attachments=attachments,
    )

    attempts = max(1, int(smtp.retries))
    for attempt in range(1, attempts + 1):
        _throttle_send_attempts(float(smtp.min_interval_sec))

        try:
            context = ssl.create_default_context()
            if int(smtp.port) == 465:
                with smtplib.SMTP_SSL(
                    smtp.host,
                    int(smtp.port),
                    context=context,
                    timeout=int(smtp.timeout_sec),
                ) as server:
                    if smtp.user and smtp.password:
                        server.login(smtp.user, smtp.password)
                    server.sendmail(from_addr, recipients, msg.as_string())
                return

            with smtplib.SMTP(
                smtp.host,
                int(smtp.port),
                timeout=int(smtp.timeout_sec),
            ) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                if smtp.user and smtp.password:
                    server.login(smtp.user, smtp.password)
                server.sendmail(from_addr, recipients, msg.as_string())
            return
        except smtplib.SMTPAuthenticationError as exc:
            raise SmtpAuthError(str(exc)) from exc
        except smtplib.SMTPRecipientsRefused as exc:
            is_transient = _is_transient_error(exc)
            if is_transient and attempt < attempts:
                _sleep_backoff(smtp, attempt)
                continue
            if is_transient:
                raise SmtpTransientError(str(exc)) from exc
            raise SmtpRecipientRefusedError(str(exc)) from exc
        except smtplib.SMTPException as exc:
            code = _smtp_status_code(exc)
            if 500 <= code < 600:
                raise SmtpPermanentError(str(exc)) from exc

            is_transient = _is_transient_error(exc)
            if is_transient and attempt < attempts:
                _sleep_backoff(smtp, attempt)
                continue
            if is_transient:
                raise SmtpTransientError(str(exc)) from exc
            raise EmailSendError(str(exc)) from exc
        except Exception as exc:
            is_transient = _is_transient_error(exc)
            if is_transient and attempt < attempts:
                _sleep_backoff(smtp, attempt)
                continue
            if is_transient:
                raise SmtpTransientError(str(exc)) from exc
            raise EmailSendError(str(exc)) from exc
