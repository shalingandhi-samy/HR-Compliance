"""Generate a rich interactive static HTML report.

All data is embedded as JSON. No server required.
Fully interactive: tabs, manager drill-downs, search, sort, charts.

Usage:
    uv run python export_report.py
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from data import load_data, STATUS_ORDER
from attendance_data import load_attendance, EXCEPTION_LABELS
from checkin_data import load_checkins
from points_data import load_points
from pto_data import load_pto

OUT_PATH = Path(__file__).parent / "hr_compliance_report.html"
GENERATED_AT = datetime.now().strftime("%B %d, %Y at %I:%M %p")


def _serialize_data() -> str:
    """Serialize all records to a single JSON blob for embedding in HTML."""
    cbls = [
        {
            "associate": r.associate,
            "win": r.win,
            "manager": r.manager,
            "status": r.status,
            "item": r.item_name,
            "due": r.due_date.strftime("%m/%d/%Y") if r.due_date else "",
            "shift": r.shift,
            "job": r.job_description,
        }
        for r in load_data()
    ]
    att = [
        {
            "associate": r.associate,
            "manager": r.manager,
            "type": EXCEPTION_LABELS.get(r.exception_type, r.exception_type),
            "date": r.exception_date.strftime("%m/%d/%Y") if r.exception_date else "",
            "shift": r.shift,
            "occ": r.occurrence or 0,
        }
        for r in load_attendance()
    ]
    chk = [
        {
            "associate": r.associate,
            "manager": r.manager,
            "needed": r.checkins_needed,
        }
        for r in load_checkins()
    ]
    pts = [
        {
            "associate": r.associate,
            "win": r.win,
            "manager": r.manager,
            "occ": r.occurrences,
            "shift": r.shift,
        }
        for r in load_points()
    ]
    pto = [
        {
            "associate": r.associate,
            "manager": r.manager,
            "submitted": r.date_submitted,
            "requested": r.date_requested,
            "position": r.position,
            "time": r.time_requested,
            "shift": r.shift,
        }
        for r in load_pto()
    ]
    return json.dumps({
        "generated": GENERATED_AT,
        "statusOrder": STATUS_ORDER,
        "cbls": cbls,
        "att": att,
        "chk": chk,
        "pts": pts,
        "pto": pto,
    }, ensure_ascii=False)


def _read_template() -> str:
    return (Path(__file__).parent / "report_template.html").read_text(encoding="utf-8")


def generate() -> str:
    data_json = _serialize_data()
    template = _read_template()
    return template.replace("__DATA_JSON__", data_json).replace("__GENERATED_AT__", GENERATED_AT)


if __name__ == "__main__":
    print("Loading data...")
    html = generate()
    OUT_PATH.write_text(html, encoding="utf-8")
    size_kb = OUT_PATH.stat().st_size // 1024
    print(f"Report saved: {OUT_PATH} ({size_kb} KB)")
