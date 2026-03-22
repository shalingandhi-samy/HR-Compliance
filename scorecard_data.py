"""Scorecard aggregation - combines all compliance modules per manager."""
from __future__ import annotations

from data import CBLRecord, STATUS_ORDER
from attendance_data import AttendanceRecord, EXCEPTION_LABELS
from checkin_data import CheckInRecord
from points_data import PointsRecord
from pto_data import PTORecord


def _all_managers(
    cbl: list[CBLRecord],
    att: list[AttendanceRecord],
    chk: list[CheckInRecord],
    pts: list[PointsRecord],
    pto: list[PTORecord],
) -> list[str]:
    """Collect every unique manager name across all five modules."""
    managers: set[str] = set()
    for r in cbl:
        managers.add(r.manager)
    for r in att:
        managers.add(r.manager)
    for r in chk:
        managers.add(r.manager)
    for r in pts:
        managers.add(r.manager)
    for r in pto:
        managers.add(r.manager)
    # Filter out blank/None manager names that sneak in from bad Excel rows
    return sorted(m for m in managers if m and m.strip())


def get_scorecard_summary(
    cbl: list[CBLRecord],
    att: list[AttendanceRecord],
    chk: list[CheckInRecord],
    pts: list[PointsRecord],
    pto: list[PTORecord],
) -> dict:
    """Build the overview table — one row per manager, all 5 metrics."""
    managers = _all_managers(cbl, att, chk, pts, pto)

    rows = []
    for m in managers:
        cbl_total = sum(1 for r in cbl if r.manager == m)
        att_total = sum(1 for r in att if r.manager == m)
        chk_total = sum(r.checkins_needed for r in chk if r.manager == m)
        pts_total = sum(1 for r in pts if r.manager == m)
        pto_total = sum(1 for r in pto if r.manager == m)
        issue_total = cbl_total + att_total + chk_total + pts_total

        rows.append({
            "manager": m,
            "cbl": cbl_total,
            "att": att_total,
            "chk": chk_total,
            "pts": pts_total,
            "pto": pto_total,
            "issue_total": issue_total,
        })

    # Sort by total compliance issues descending
    rows.sort(key=lambda r: r["issue_total"], reverse=True)
    return {"rows": rows}


def get_manager_scorecard(
    manager: str,
    cbl: list[CBLRecord],
    att: list[AttendanceRecord],
    chk: list[CheckInRecord],
    pts: list[PointsRecord],
    pto: list[PTORecord],
) -> dict:
    """Full compliance scorecard for a single manager."""
    # --- CBLs ---
    cbl_f = [r for r in cbl if r.manager == manager]
    cbl_by_status = {s: 0 for s in STATUS_ORDER}
    for r in cbl_f:
        cbl_by_status[r.status] += 1

    # --- Attendance exceptions ---
    att_f = [r for r in att if r.manager == manager]
    att_types = sorted({r.exception_type for r in att_f})
    att_by_type: dict[str, int] = {}
    for r in att_f:
        att_by_type[r.exception_type] = att_by_type.get(r.exception_type, 0) + 1

    # group by associate for the detail table
    att_by_assoc: dict[str, list[AttendanceRecord]] = {}
    for r in att_f:
        att_by_assoc.setdefault(r.associate, []).append(r)
    att_sorted_assocs = sorted(
        att_by_assoc.keys(), key=lambda a: len(att_by_assoc[a]), reverse=True
    )

    # --- Check-ins ---
    chk_f = sorted(
        [r for r in chk if r.manager == manager],
        key=lambda r: r.checkins_needed,
        reverse=True,
    )
    chk_total = sum(r.checkins_needed for r in chk_f)

    # --- Points 5+ ---
    pts_f = sorted(
        [r for r in pts if r.manager == manager],
        key=lambda r: r.occurrences,
        reverse=True,
    )

    # --- PTO ---
    pto_f = [r for r in pto if r.manager == manager]

    return {
        "manager": manager,
        "cbl": {
            "total": len(cbl_f),
            "by_status": cbl_by_status,
            "records": sorted(cbl_f, key=lambda r: STATUS_ORDER.index(r.status)),
        },
        "attendance": {
            "total": len(att_f),
            "by_type": att_by_type,
            "all_types": att_types,
            "by_assoc": att_by_assoc,
            "sorted_assocs": att_sorted_assocs,
        },
        "checkins": {
            "total": chk_total,
            "associates": len(chk_f),
            "records": chk_f,
        },
        "points": {
            "total": len(pts_f),
            "records": pts_f,
        },
        "pto": {
            "total": len(pto_f),
            "records": pto_f,
        },
        "status_order": STATUS_ORDER,
        "exception_labels": EXCEPTION_LABELS,
    }
