"""Aggregate compliance totals by shift across all modules."""
from __future__ import annotations

from data import CBLRecord
from attendance_data import AttendanceRecord
from points_data import PointsRecord
from pto_data import PTORecord


def get_shift_breakdown(
    cbl: list[CBLRecord],
    att: list[AttendanceRecord],
    pts: list[PointsRecord],
    pto: list[PTORecord],
) -> dict:
    """Return per-shift totals for each compliance module."""
    shifts: set[str] = set()
    for r in cbl: shifts.add(r.shift)
    for r in att: shifts.add(r.shift)
    for r in pts: shifts.add(r.shift)
    for r in pto: shifts.add(r.shift)
    sorted_shifts = sorted(shifts)

    rows = []
    for s in sorted_shifts:
        cbl_t = sum(1 for r in cbl if r.shift == s)
        att_t = sum(1 for r in att if r.shift == s)
        pts_t = sum(1 for r in pts if r.shift == s)
        pto_t = sum(1 for r in pto if r.shift == s)
        rows.append({
            "shift": s,
            "cbl": cbl_t,
            "att": att_t,
            "pts": pts_t,
            "pto": pto_t,
            "total": cbl_t + att_t + pts_t + pto_t,
        })

    rows.sort(key=lambda r: r["total"], reverse=True)
    return {
        "shifts": sorted_shifts,
        "rows": rows,
    }
