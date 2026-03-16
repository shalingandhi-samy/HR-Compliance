"""Data loading and processing for Pending Time Off Requests tab."""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime, time
from typing import Optional

import openpyxl

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "phl5_compliance.xlsx")


def _fmt_duration(t: time) -> str:
    """Format a time object as hours/minutes string."""
    if t is None:
        return "—"
    total_mins = t.hour * 60 + t.minute
    if total_mins == 0:
        return "Full Day"
    if t.minute == 0:
        return f"{t.hour}h"
    if t.hour == 0:
        return f"{t.minute}m"
    return f"{t.hour}h {t.minute}m"


@dataclass
class PTORecord:
    date_submitted: str
    date_requested: str
    win: str
    associate: str
    team: str
    position: str
    time_requested: str
    manager: str
    shift: str


_cache: list[PTORecord] = []
_loaded_at: Optional[datetime] = None


def load_pto(force: bool = False) -> list[PTORecord]:
    """Load Pending Time Off Requests sheet, caching in memory."""
    global _cache, _loaded_at
    if _cache and not force:
        return _cache

    records: list[PTORecord] = []
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True, read_only=True)
    ws = wb["Pending Time Off Requests"]

    header_passed = False
    for row in ws.iter_rows(values_only=True):
        if not header_passed:
            if row[0] == "Date Submitted":
                header_passed = True
            continue

        if not row[0]:
            continue

        # Date submitted — strip the newline/time portion
        date_sub = str(row[0]).split("\n")[0].strip() if row[0] else ""

        # Date requested — format as MM/DD/YYYY
        if row[1] and isinstance(row[1], datetime):
            date_req = row[1].strftime("%m/%d/%Y")
        else:
            date_req = str(row[1]).strip() if row[1] else ""

        # Shift — strip " (United States of America)"
        shift_raw = str(row[8]).strip() if row[8] else "Unknown"
        shift = shift_raw.split(" (")[0].strip()

        records.append(PTORecord(
            date_submitted=date_sub,
            date_requested=date_req,
            win=str(row[2]).strip() if row[2] else "",
            associate=str(row[3]).strip() if row[3] else "",
            team=str(row[4]).strip() if row[4] else "",
            position=str(row[5]).strip() if row[5] else "",
            time_requested=_fmt_duration(row[6]) if isinstance(row[6], time) else "Full Day",
            manager=str(row[7]).strip() if row[7] else "No Manager",
            shift=shift,
        ))

    wb.close()
    _cache = sorted(records, key=lambda r: r.date_requested)
    _loaded_at = datetime.now()
    return _cache


def get_pto_summary(records: list[PTORecord]) -> dict:
    """Build summary stats for the PTO dashboard."""
    manager_stats: dict[str, int] = {}
    shift_stats: dict[str, int] = {}
    position_stats: dict[str, int] = {}

    for r in records:
        manager_stats[r.manager] = manager_stats.get(r.manager, 0) + 1
        shift_stats[r.shift] = shift_stats.get(r.shift, 0) + 1
        position_stats[r.position] = position_stats.get(r.position, 0) + 1

    sorted_managers = sorted(manager_stats, key=lambda m: manager_stats[m], reverse=True)
    sorted_shifts = sorted(shift_stats, key=lambda s: shift_stats[s], reverse=True)

    return {
        "total": len(records),
        "manager_stats": manager_stats,
        "sorted_managers": sorted_managers,
        "shift_stats": shift_stats,
        "sorted_shifts": sorted_shifts,
        "position_stats": position_stats,
    }


def get_pto_manager_detail(records: list[PTORecord], manager: str) -> dict:
    """Get all PTO requests under a manager."""
    filtered = [r for r in records if r.manager == manager]
    return {
        "manager": manager,
        "total": len(filtered),
        "records": filtered,
    }