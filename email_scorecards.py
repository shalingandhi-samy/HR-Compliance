"""Email manager scorecards — sends via Outlook COM (Windows) with SMTP fallback.

Usage:
    uv run python email_scorecards.py           # send all
    uv run python email_scorecards.py --dry-run  # preview only
"""
from __future__ import annotations

import configparser
import logging
import re
import smtplib
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from data import load_data
from attendance_data import load_attendance
from checkin_data import load_checkins
from points_data import load_points
from pto_data import load_pto
from scorecard_data import get_manager_scorecard, get_scorecard_summary

logger = logging.getLogger(__name__)

CFG_PATH = Path(__file__).parent / "manager_emails.cfg"


def _load_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    cfg.read(CFG_PATH)
    return cfg


def _derive_email(name: str) -> str:
    """Turn 'John Smith' into 'john.smith@wal-mart.com'."""
    clean = re.sub(r"[^a-zA-Z\s-]", "", name).strip().lower()
    parts = clean.split()
    if len(parts) >= 2:
        return f"{parts[0]}.{parts[-1]}@wal-mart.com"
    return f"{clean}@wal-mart.com"


def get_manager_email(name: str, cfg: configparser.ConfigParser) -> str:
    override = cfg.get("emails", name, fallback="")
    return override.strip() if override.strip() else _derive_email(name)


def _send_via_outlook(to_email: str, subject: str, html_body: str) -> None:
    """Send via local Outlook COM — uses your already-logged-in Outlook session."""
    try:
        import win32com.client as win32  # type: ignore
    except ImportError:
        raise RuntimeError("pywin32 not installed")
    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = to_email
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Send()
        logger.info(f"Sent via Outlook COM -> {to_email}")
    except Exception as exc:
        raise RuntimeError(f"Outlook COM error: {exc}") from exc


def _send_via_smtp(to_email: str, subject: str, html_body: str, cfg: configparser.ConfigParser) -> None:
    """Fallback: send via configured SMTP relay."""
    smtp_host = cfg.get("settings", "smtp_host", fallback="smtp.office365.com")
    smtp_port = int(cfg.get("settings", "smtp_port", fallback="587"))
    from_addr = cfg.get("settings", "from_address", fallback="")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = from_addr
    msg["To"]      = to_email
    msg.attach(MIMEText(html_body, "html"))
    with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
        server.starttls()
        server.sendmail(from_addr, [to_email], msg.as_string())
    logger.info(f"Sent via SMTP -> {to_email}")


def send_email(to_email: str, subject: str, html_body: str) -> None:
    """Try Outlook COM first, fall back to SMTP."""
    cfg = _load_config()
    try:
        _send_via_outlook(to_email, subject, html_body)
    except RuntimeError as exc:
        logger.warning(f"Outlook COM unavailable ({exc}), trying SMTP...")
        _send_via_smtp(to_email, subject, html_body, cfg)


def _render_scorecard_email(manager: str, data: dict) -> str:
    """Generate a clean HTML email body for one manager's scorecard."""
    now  = datetime.now().strftime("%B %d, %Y")
    cbl  = data["cbl"]
    att  = data["attendance"]
    chk  = data["checkins"]
    pts  = data["points"]
    pto  = data["pto"]

    def metric_row(label: str, value: int, color: str) -> str:
        flag  = "&#9888;" if value > 0 else "&#10003;"
        style = f"color:{color};font-weight:bold" if value > 0 else "color:#2a8703"
        return (
            f'<tr><td style="padding:8px 16px;border-bottom:1px solid #f0f0f0">{label}</td>'
            f'<td style="padding:8px 16px;border-bottom:1px solid #f0f0f0;text-align:center;{style}">{flag} {value}</td></tr>'
        )

    total        = cbl["total"] + att["total"] + chk["total"] + pts["total"]
    banner_color = "#2a8703" if total == 0 else "#ea1100" if total > 10 else "#995213"
    status_msg   = "All Clear! &#127881;" if total == 0 else f"{total} Open Items Require Attention"

    return f"""
    <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8f9fa">
      <div style="background:#0053e2;padding:24px 32px">
        <div style="color:white;font-size:22px;font-weight:900;letter-spacing:1px">PHL5 COMPLIANCE</div>
        <div style="color:rgba(255,255,255,0.8);font-size:13px;margin-top:4px">Manager Scorecard &mdash; {now}</div>
      </div>
      <div style="background:{banner_color};padding:14px 32px;color:white;font-weight:bold;font-size:15px">
        {manager} &mdash; {status_msg}
      </div>
      <div style="background:white;padding:24px 32px">
        <table style="width:100%;border-collapse:collapse;font-size:14px">
          <thead>
            <tr style="background:#041f41;color:white">
              <th style="padding:10px 16px;text-align:left">Metric</th>
              <th style="padding:10px 16px;text-align:center">Open Items</th>
            </tr>
          </thead>
          <tbody>
            {metric_row('&#128218; CBL Pending', cbl['total'], '#0053e2')}
            {metric_row('&#128197; Attendance Exceptions', att['total'], '#995213')}
            {metric_row('&#9989; Check-Ins Needed', chk['total'], '#ea1100')}
            {metric_row('&#9888; Associates at Points 5+', pts['total'], '#7f1d1d')}
            {metric_row('&#128203; Time Off Requests', pto['total'], '#2a8703')}
          </tbody>
        </table>
      </div>
      <div style="background:#f0f4ff;padding:16px 32px;border-top:3px solid #0053e2">
        <p style="font-size:12px;color:#555;margin:0">
          Auto-generated by the PHL5 Compliance Dashboard. Do not reply.
        </p>
      </div>
    </div>
    """


def send_all_scorecards(dry_run: bool = False) -> dict:
    """Send scorecard emails to all managers."""
    cfg      = _load_config()
    cbl      = load_data()
    att      = load_attendance()
    chk      = load_checkins()
    pts      = load_points()
    pto      = load_pto()
    summary  = get_scorecard_summary(cbl, att, chk, pts, pto)
    managers = [r["manager"] for r in summary["rows"]]

    sent, failed, skipped = [], [], []
    for manager in managers:
        email   = get_manager_email(manager, cfg)
        data    = get_manager_scorecard(manager, cbl, att, chk, pts, pto)
        html    = _render_scorecard_email(manager, data)
        subject = f"PHL5 Compliance Scorecard \u2014 {manager} \u2014 {datetime.now().strftime('%b %d, %Y')}"
        if dry_run:
            logger.info(f"[DRY RUN] Would email {manager} -> {email}")
            skipped.append({"manager": manager, "email": email})
            continue
        try:
            send_email(email, subject, html)
            sent.append({"manager": manager, "email": email})
        except Exception as exc:
            logger.warning(f"Failed to email {manager}: {exc}")
            failed.append({"manager": manager, "email": email, "error": str(exc)})

    return {"sent": sent, "failed": failed, "skipped": skipped, "dry_run": dry_run}


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    dry    = "--dry-run" in sys.argv
    result = send_all_scorecards(dry_run=dry)
    print(f"\nSent: {len(result['sent'])} | Failed: {len(result['failed'])} | Skipped: {len(result['skipped'])}")
