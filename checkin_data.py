"""Data loading and processing for Manager Check Ins tab."""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

import openpyxl

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "phl5_compliance.xlsx")

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
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True, read_only=True)
    ws = wb["Manager Check Ins"]

    # First two rows are title/sub-header — skip them
    data_started = False
    for row in ws.iter_rows(values_only=True):
        if not data_started:
            # Header row has 'Associates' in col 0
            if row[0] == "Associates":
                data_started = True
            continue

        if not row[0]:  # skip empty rows
            continue

        try:
            checkins = int(row[2]) if row[2] is not None else 0
        except (ValueError, TypeError):
            checkins = 0

        records.append(CheckInRecord(
            associate=str(row[0]).strip(),
            manager=str(row[1]).strip() if row[1] else "No Manager",
            checkins_needed=checkins,
            status=str(row[3]).strip() if row[3] else "LATE",
        ))

    wb.close()
    _cache = records
    _loaded_at = datetime.now()
    return records


def get_checkin_summary(records: list[CheckInRecord]) -> dict:
    """Build summary stats for the check-in dashboard."""
    total = len(records)

    # Bucket totals (how many associates need 1/2/3/4 check-ins)
    bucket_totals = {b: 0 for b in CHECKIN_BUCKETS}
    for r in records:
        key = min(r.checkins_needed, 4)  # cap at 4
        if key in bucket_totals:
            bucket_totals[key] += 1

    # Per manager per bucket
    manager_stats: dict[str, dict[int, int]] = {}
    for r in records:
        if r.manager not in manager_stats:
            manager_stats[r.manager] = {b: 0 for b in CHECKIN_BUCKETS}
        key = min(r.checkins_needed, 4)
        if key in manager_stats[r.manager]:
            manager_stats[r.manager][key] += 1

    manager_totals = {m: sum(v.values()) for m, v in manager_stats.items()}
    sorted_managers = sorted(
        manager_stats.keys(), key=lambda m: manager_totals[m], reverse=True
    )

    return {
        "total": total,
        "bucket_totals": bucket_totals,
        "manager_stats": manager_stats,
        "manager_totals": manager_totals,
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

    return {
        "manager": manager,
        "total": len(filtered),
        "records": sorted_assocs,
        "bucket_totals": bucket_totals,
        "buckets": CHECKIN_BUCKETS,
    }