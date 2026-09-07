"""Data loading and processing for Manager Check Ins tab."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional

import openpyxl

from onedrive_client import get_workbook
from sheet_utils import col, find_header_row

# Number of check-ins needed breakdown
CHECKIN_BUCKETS = [1, 2, 3, 4]


@dataclass
class CheckInRecord:
    associate: str
    manager: str
    checkins_needed: int
    status: str  # e.g. 'LATE'


_cache: list[CheckInRecord] = []
_loaded_at: Optional[datetime] = None


def load_checkins(force: bool = False) -> list[CheckInRecord]:
    """Load Manager Check Ins sheet, caching in memory."""
    global _cache, _loaded_at
    if _cache and not force:
        return _cache

    records: list[CheckInRecord] = []
    wb = get_workbook()
    ws = wb["Manager Check-Ins"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Header row: Associate, Manager, Date Hired, Check-Ins Needed,
    # FORMULA Do Not Modify (status), WIN, SHIFT -- located by name so a
    # future rename/reorder (like the 'Associates'->'Associate' rename and
    # 'Date Hired' insertion that broke this before) fails loudly instead
    # of silently returning 0.
    header_idx, hmap = find_header_row(rows, ["Associate", "Associates"])
    c_associate = col(hmap, "Associate", "Associates")
    c_manager = col(hmap, "Manager")
    c_checkins = col(hmap, "Check-Ins Needed", "Checkins Needed")
    c_status = col(hmap, "FORMULA Do Not Modify", required=False)

    for row in rows[header_idx + 1:]:
        if not row[c_associate]:  # skip empty rows
            continue

        try:
            checkins = int(row[c_checkins]) if row[c_checkins] is not None else 0
        except (ValueError, TypeError):
            checkins = 0

        status_val = row[c_status] if c_status is not None else None
        records.append(CheckInRecord(
            associate=str(row[c_associate]).strip(),
            manager=str(row[c_manager]).strip() if row[c_manager] else "No Manager",
            checkins_needed=checkins,
            status=str(status_val).strip() if status_val else "LATE",
        ))

    _cache = records
    _loaded_at = datetime.now()
    return records


def get_checkin_summary(records: list[CheckInRecord]) -> dict:
    """Build summary stats for the check-in dashboard."""
    # Grand total = sum of all check-ins needed across all associates
    total_checkins = sum(r.checkins_needed for r in records)
    total_associates = len(records)

    # Bucket totals (how many associates need 1/2/3/4 check-ins)
    bucket_totals = {b: 0 for b in CHECKIN_BUCKETS}
    for r in records:
        key = min(r.checkins_needed, 4)
        if key in bucket_totals:
            bucket_totals[key] += 1

    # Per manager: count associates per bucket + sum of checkins_needed
    manager_stats: dict[str, dict[int, int]] = {}
    manager_checkin_sums: dict[str, int] = {}
    for r in records:
        if r.manager not in manager_stats:
            manager_stats[r.manager] = {b: 0 for b in CHECKIN_BUCKETS}
            manager_checkin_sums[r.manager] = 0
        key = min(r.checkins_needed, 4)
        if key in manager_stats[r.manager]:
            manager_stats[r.manager][key] += 1
        manager_checkin_sums[r.manager] += r.checkins_needed

    # Sort by total check-ins needed desc
    sorted_managers = sorted(
        manager_stats.keys(),
        key=lambda m: manager_checkin_sums[m],
        reverse=True,
    )

    return {
        "total": total_checkins,
        "total_associates": total_associates,
        "bucket_totals": bucket_totals,
        "manager_stats": manager_stats,
        "manager_totals": manager_checkin_sums,
        "sorted_managers": sorted_managers,
        "buckets": CHECKIN_BUCKETS,
    }


def get_checkin_manager_detail(records: list[CheckInRecord], manager: str) -> dict:
    """Get all associates under a manager with their check-in counts."""
    filtered = [r for r in records if r.manager == manager]

    # Sort by check-ins needed desc (most urgent first)
    sorted_assocs = sorted(filtered, key=lambda r: r.checkins_needed, reverse=True)

    bucket_totals = {b: 0 for b in CHECKIN_BUCKETS}
    for r in filtered:
        key = min(r.checkins_needed, 4)
        if key in bucket_totals:
            bucket_totals[key] += 1

    total_checkins = sum(r.checkins_needed for r in filtered)

    return {
        "manager": manager,
        "total": total_checkins,
        "total_associates": len(filtered),
        "records": sorted_assocs,
        "bucket_totals": bucket_totals,
        "buckets": CHECKIN_BUCKETS,
    }