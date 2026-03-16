"""Data loading and processing for Attendance Points 5+ tab."""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

import openpyxl

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "phl5_compliance.xlsx")


@dataclass
class PointsRecord:
    win: str
    associate: str
    code: str
    occurrences: float
    manager: str
    shift: str


_cache: list[PointsRecord] = []
_loaded_at: Optional[datetime] = None


def load_points(force: bool = False) -> list[PointsRecord]:
    """Load Attendance Points 5+ sheet, caching in memory."""
    global _cache, _loaded_at
    if _cache and not force:
        return _cache

    records: list[PointsRecord] = []
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True, read_only=True)
    ws = wb["Attendance Points 5+"]

    header_passed = False
    for row in ws.iter_rows(values_only=True):
        if not header_passed:
            if row[0] == "Associate WIN":
                header_passed = True
            continue

        if not row[0]:
            continue

        try:
            occ = float(row[3]) if row[3] is not None else 0.0
        except (ValueError, TypeError):
            occ = 0.0

        shift_raw = str(row[5]).strip() if row[5] else "Unknown"
        shift = shift_raw.split(" (")[0].strip()

        records.append(PointsRecord(
            win=str(row[0]).strip(),
            associate=str(row[1]).strip() if row[1] else "",
            code=str(row[2]).strip() if row[2] else "",
            occurrences=occ,
            manager=str(row[4]).strip() if row[4] else "No Manager",
            shift=shift,
        ))

    wb.close()
    _cache = sorted(records, key=lambda r: r.occurrences, reverse=True)
    _loaded_at = datetime.now()
    return _cache


def get_points_summary(records: list[PointsRecord]) -> dict:
    """Build summary stats for the points dashboard."""
    manager_stats: dict[str, dict] = {}
    for r in records:
        if r.manager not in manager_stats:
            manager_stats[r.manager] = {"total": 0, "max_occ": 0.0}
        manager_stats[r.manager]["total"] += 1
        if r.occurrences > manager_stats[r.manager]["max_occ"]:
            manager_stats[r.manager]["max_occ"] = r.occurrences

    sorted_managers = sorted(
        manager_stats.keys(),
        key=lambda m: manager_stats[m]["total"],
        reverse=True,
    )

    shift_stats: dict[str, int] = {}
    for r in records:
        shift_stats[r.shift] = shift_stats.get(r.shift, 0) + 1

    return {
        "total": len(records),
        "manager_stats": manager_stats,
        "sorted_managers": sorted_managers,
        "shift_stats": shift_stats,
    }


def get_points_manager_detail(records: list[PointsRecord], manager: str) -> dict:
    """Get all associates under a manager."""
    filtered = sorted(
        [r for r in records if r.manager == manager],
        key=lambda r: r.occurrences,
        reverse=True,
    )
    return {
        "manager": manager,
        "total": len(filtered),
        "records": filtered,
    }