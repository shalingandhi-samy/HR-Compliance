"""Data loading and processing for Attendance Exceptions tab."""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime, time
from typing import Optional

import openpyxl

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "phl5_compliance.xlsx")

EXCEPTION_TYPES = [
    "AT_ABSENT",
    "AT_ABSENT_NO_CALL",
    "AT_EARLY_IN",
    "AT_EARLY_OUT",
    "AT_INC_SHIFT",
    "AT_LATE_IN",
]

EXCEPTION_LABELS = {
    "AT_ABSENT": "Absent",
    "AT_ABSENT_NO_CALL": "Absent No Call",
    "AT_EARLY_IN": "Early In",
    "AT_EARLY_OUT": "Early Out",
    "AT_INC_SHIFT": "Incomplete Shift",
    "AT_LATE_IN": "Late In",
}


@dataclass
class AttendanceRecord:
    action: str
    exception_date: Optional[datetime]
    days_before_auth: Optional[int]
    win: Optional[int]
    associate: str
    occurrence: Optional[float]
    team: str
    team_desc: str
    exception_type: str
    duration: Optional[time]
    manager: str
    shift: str


_cache: list[AttendanceRecord] = []
_loaded_at: Optional[datetime] = None


def _parse_shift(raw: Optional[str]) -> str:
    if not raw:
        return "Unknown"
    parts = str(raw).split(" - ", 1)
    if len(parts) == 2:
        num = parts[0].strip()
        name = parts[1].split(" (")[0].strip()
        return f"Shift {num} - {name}"
    return str(raw)


def load_attendance(force: bool = False) -> list[AttendanceRecord]:
    """Load Attendance Exceptions sheet, caching in memory."""
    global _cache, _loaded_at
    if _cache and not force:
        return _cache

    records: list[AttendanceRecord] = []
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True, read_only=True)
    ws = wb["Attendance Exceptions"]

    header_found = False
    for row in ws.iter_rows(values_only=True):
        if not header_found:
            if row[0] == "Action":
                header_found = True
            continue

        # Only process Authorize rows (they have the actual data)
        if not row[0] or str(row[0]).strip() != "Authorize":
            continue
        if not row[4]:  # skip if no associate name
            continue

        records.append(AttendanceRecord(
            action=str(row[0]).strip(),
            exception_date=row[1] if isinstance(row[1], datetime) else None,
            days_before_auth=row[2],
            win=row[3],
            associate=str(row[4]).strip(),
            occurrence=row[5],
            team=str(row[6]).strip() if row[6] else "",
            team_desc=str(row[7]).strip() if row[7] else "",
            exception_type=str(row[8]).strip() if row[8] else "UNKNOWN",
            duration=row[9] if isinstance(row[9], time) else None,
            manager=str(row[10]).strip() if row[10] else "No Manager",
            shift=_parse_shift(row[11]),
        ))

    wb.close()
    _cache = records
    _loaded_at = datetime.now()
    return records


def get_attendance_summary(records: list[AttendanceRecord]) -> dict:
    """Build summary stats for the attendance dashboard."""
    total = len(records)

    # Collect all exception types actually present
    all_types = sorted({r.exception_type for r in records})

    # Per manager per exception type
    manager_stats: dict[str, dict[str, int]] = {}
    for r in records:
        if r.manager not in manager_stats:
            manager_stats[r.manager] = {t: 0 for t in all_types}
        manager_stats[r.manager][r.exception_type] = manager_stats[r.manager].get(r.exception_type, 0) + 1
    # Per shift per exception type
    shift_stats: dict[str, dict[str, int]] = {}
    for r in records:
        if r.shift not in shift_stats:
            shift_stats[r.shift] = {t: 0 for t in all_types}
        shift_stats[r.shift][r.exception_type] = shift_stats[r.shift].get(r.exception_type, 0) + 1
    # Type totals
    type_totals = {t: sum(manager_stats[m].get(t, 0) for m in manager_stats) for t in all_types}

    # Sort managers by total desc
    manager_totals = {m: sum(v.values()) for m, v in manager_stats.items()}
    sorted_managers = sorted(manager_stats.keys(), key=lambda m: manager_totals[m], reverse=True)
    sorted_shifts = sorted(shift_stats.keys())

    return {
        "total": total,
        "all_types": all_types,
        "type_totals": type_totals,
        "manager_stats": manager_stats,
        "manager_totals": manager_totals,
        "sorted_managers": sorted_managers,
        "shift_stats": shift_stats,
        "sorted_shifts": sorted_shifts,
    }


def get_attendance_manager_detail(records: list[AttendanceRecord], manager: str) -> dict:
    """Get all associates + their exceptions for a specific manager."""
    filtered = [r for r in records if r.manager == manager]
    all_types = sorted({r.exception_type for r in filtered})

    assoc_map: dict[str, list[AttendanceRecord]] = {}
    for r in filtered:
        if r.associate not in assoc_map:
            assoc_map[r.associate] = []
        assoc_map[r.associate].append(r)

    sorted_assocs = sorted(
        assoc_map.keys(),
        key=lambda a: len(assoc_map[a]),
        reverse=True,
    )

    return {
        "manager": manager,
        "total": len(filtered),
        "all_types": all_types,
        "assoc_map": assoc_map,
        "sorted_assocs": sorted_assocs,
    }