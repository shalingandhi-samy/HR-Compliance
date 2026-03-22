"""Associate Lookup — search one associate across all 5 compliance modules."""
from __future__ import annotations

from data import CBLRecord
from attendance_data import AttendanceRecord, EXCEPTION_LABELS
from checkin_data import CheckInRecord
from points_data import PointsRecord
from pto_data import PTORecord


def search_associate(
    query: str,
    cbl: list[CBLRecord],
    att: list[AttendanceRecord],
    chk: list[CheckInRecord],
    pts: list[PointsRecord],
    pto: list[PTORecord],
) -> dict:
    """Return all compliance records matching the associate name query."""
    q = query.strip().lower()
    if not q:
        return {"query": query, "found": False, "results": []}

    # Find all matching associate names across modules
    candidates: set[str] = set()
    for r in cbl:
        if q in r.associate.lower():
            candidates.add(r.associate)
    for r in att:
        if q in r.associate.lower():
            candidates.add(r.associate)
    for r in chk:
        if q in r.associate.lower():
            candidates.add(r.associate)
    for r in pts:
        if q in r.associate.lower():
            candidates.add(r.associate)
    for r in pto:
        if q in r.associate.lower():
            candidates.add(r.associate)

    if not candidates:
        return {"query": query, "found": False, "results": []}

    results = []
    for name in sorted(candidates):
        cbl_r = [r for r in cbl if r.associate == name]
        att_r = [r for r in att if r.associate == name]
        chk_r = [r for r in chk if r.associate == name]
        pts_r = [r for r in pts if r.associate == name]
        pto_r = [r for r in pto if r.associate == name]

        # Derive manager from first match found
        manager = (
            (cbl_r[0].manager if cbl_r else None)
            or (att_r[0].manager if att_r else None)
            or (chk_r[0].manager if chk_r else None)
            or (pts_r[0].manager if pts_r else None)
            or (pto_r[0].manager if pto_r else None)
            or "Unknown"
        )
        shift = (
            (cbl_r[0].shift if cbl_r else None)
            or (att_r[0].shift if att_r else None)
            or (pts_r[0].shift if pts_r else None)
            or "Unknown"
        )

        results.append({
            "name": name,
            "manager": manager,
            "shift": shift,
            "cbl": cbl_r,
            "attendance": att_r,
            "checkins": chk_r,
            "points": pts_r,
            "pto": pto_r,
            "total_items": len(cbl_r) + len(att_r) + len(chk_r) + len(pts_r) + len(pto_r),
        })

    results.sort(key=lambda r: r["total_items"], reverse=True)
    return {
        "query": query,
        "found": True,
        "count": len(results),
        "results": results,
        "exception_labels": EXCEPTION_LABELS,
    }
