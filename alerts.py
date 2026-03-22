"""Compliance alerts — sends Teams webhook notifications when thresholds are exceeded.

Configuration (set as environment variables or edit THRESHOLDS below):
    TEAMS_WEBHOOK_URL   — incoming webhook URL from your Teams channel
    ALERT_CBL_THRESHOLD — CBL total that triggers an alert (default: 50)
    ALERT_ATT_THRESHOLD — Attendance exceptions threshold (default: 30)
    ALERT_PTS_THRESHOLD — Points 5+ threshold (default: 20)
"""
from __future__ import annotations

import json
import logging
import os
import urllib.error
import urllib.request
from datetime import datetime

from data import CBLRecord
from attendance_data import AttendanceRecord
from checkin_data import CheckInRecord, get_checkin_summary
from points_data import PointsRecord
from pto_data import PTORecord

logger = logging.getLogger(__name__)

TEAMS_WEBHOOK_URL = os.getenv("TEAMS_WEBHOOK_URL", "")

THRESHOLDS = {
    "cbl": int(os.getenv("ALERT_CBL_THRESHOLD", "50")),
    "att": int(os.getenv("ALERT_ATT_THRESHOLD", "30")),
    "pts": int(os.getenv("ALERT_PTS_THRESHOLD", "20")),
}

PROXIES = {
    "http":  "http://sysproxy.wal-mart.com:8080",
    "https": "http://sysproxy.wal-mart.com:8080",
}


def _send_teams_message(title: str, body: str) -> None:
    """Post a simple Adaptive Card message to Teams via webhook."""
    if not TEAMS_WEBHOOK_URL:
        logger.info("[Alerts] Teams webhook not configured — skipping notification.")
        return

    payload = json.dumps({
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.4",
                "body": [
                    {"type": "TextBlock", "text": title,
                     "weight": "Bolder", "size": "Medium", "color": "Attention"},
                    {"type": "TextBlock", "text": body, "wrap": True},
                    {"type": "TextBlock",
                     "text": f"Generated: {datetime.now().strftime('%b %d %Y %I:%M %p')}",
                     "isSubtle": True, "size": "Small"},
                ],
            },
        }],
    }).encode("utf-8")

    req = urllib.request.Request(
        TEAMS_WEBHOOK_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    proxy_handler = urllib.request.ProxyHandler(PROXIES)
    opener = urllib.request.build_opener(proxy_handler)
    try:
        with opener.open(req, timeout=15) as resp:
            logger.info(f"[Alerts] Teams notification sent (HTTP {resp.status})")
    except urllib.error.URLError as exc:
        logger.warning(f"[Alerts] Teams notification failed: {exc}")


def check_and_send_alerts(
    cbl: list[CBLRecord],
    att: list[AttendanceRecord],
    chk: list[CheckInRecord],
    pts: list[PointsRecord],
    pto: list[PTORecord],
) -> None:
    """Check totals against thresholds and fire Teams alerts if breached."""
    chk_summary = get_checkin_summary(chk)
    totals = {
        "cbl": len(cbl),
        "att": len(att),
        "pts": len(pts),
    }

    breaches = []
    labels = {
        "cbl": "CBL Pending",
        "att": "Attendance Exceptions",
        "pts": "Associates at 5+ Points",
    }

    for key, threshold in THRESHOLDS.items():
        val = totals[key]
        if val >= threshold:
            breaches.append(f"\u26a0\ufe0f **{labels[key]}**: {val} (threshold: {threshold})")

    if breaches:
        body = "\n\n".join(breaches) + "\n\nPlease review the PHL5 Compliance Dashboard."
        _send_teams_message("\U0001f6a8 PHL5 Compliance Alert", body)
        logger.info(f"[Alerts] {len(breaches)} threshold(s) breached — notification sent.")
    else:
        logger.info("[Alerts] All metrics within thresholds.")
