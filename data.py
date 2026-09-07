"""Data loading and processing for PHL5 Compliance Dashboard."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Optional

import openpyxl

from onedrive_client import get_workbook
from sheet_utils import col, find_header_row

STATUS_ORDER = ["Overdue", "7 Days", "14 Days", "30 Days", "60 Days"]


@dataclass
class CBLRecord:
    associate: str
    win: Optional[int]
    user_id: str
    job_description: str
    item_name: str
    due_date: Optional[datetime]
    status: str
    manager: str
    shift: str


_cache: list[CBLRecord] = []
_loaded_at: Optional[datetime] = None


# Job Description (column D) values end in a shift code like
# 'WAREHOUSE WORKER_S5' -- the last 2 characters are the shift (S1-S7).
SHIFT_CODE_RE = re.compile(r"(S[1-7])\s*$", re.IGNORECASE)


def _parse_shift(job_description: Optional[str]) -> str:
    """Extract the shift code from the tail of the Job Description string.

    Source changed (Sept 2026): shift used to come from the Position column
    (M / row[12], e.g. '4 - Weekend (United States of America)'). It now
    comes from Job Description (D / row[3]) instead -- last 2 characters
    are the shift code, e.g. 'WAREHOUSE WORKER_S5' -> 'S5'.
    """
    if not job_description:
        return "Unknown"
    m = SHIFT_CODE_RE.search(str(job_description).strip())
    if m:
        return m.group(1).upper()
    return "Unknown"


def _status_from_due_date(due_date: Optional[datetime], today: Optional[date] = None) -> Optional[str]:
    """Compute the TRUE current status bucket from the due date, not from
    the source sheet's Late/7/14/30/60 Days flag columns.

    Why: the source sheet logs the same training requirement once per
    historical data pull, and never prunes old snapshot rows. As the due
    date approaches, each new pull gets a fresh row with the countdown
    bucket shifted closer (60 -> 30 -> 14 -> 7 -> Late), but the OLD rows
    stick around too. Trusting each row's flag independently means the
    same person's same course gets counted once per historical snapshot
    instead of once. Recomputing fresh from the due date (deduped by
    associate+WIN+item+due date beforehand) gives the true current count.
    Boundaries mirror the source sheet's own non-overlapping windows.
    """
    if not isinstance(due_date, datetime):
        return None
    if today is None:
        today = date.today()
    due = due_date.date()
    if due < today:
        return "Overdue"
    if due <= today + timedelta(days=7):
        return "7 Days"
    if due <= today + timedelta(days=14):
        return "14 Days"
    if due <= today + timedelta(days=30):
        return "30 Days"
    if due <= today + timedelta(days=60):
        return "60 Days"
    return None  # beyond the tracked horizon


def load_data(force: bool = False) -> list[CBLRecord]:
    """Load Excel data, caching in memory."""
    global _cache, _loaded_at
    if _cache and not force:
        return _cache

    wb = get_workbook()
    ws = wb["ULearns"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    header_idx, hmap = find_header_row(rows, ["Associate", "Associates"])
    c_associate = col(hmap, "Associate", "Associates")
    c_win = col(hmap, "WIN")
    c_user_id = col(hmap, "User ID")
    c_job_description = col(hmap, "Job Description")
    c_item_name = col(hmap, "Item Name")
    c_due_date = col(hmap, "Due Date")
    c_manager = col(hmap, "Manager")

    # Dedupe by the true identity of a training requirement -- the source
    # sheet logs the same (associate, course, due date) once per historical
    # pull, so raw rows massively overcount distinct pending items (in
    # practice: ~1,798 raw rows vs ~945 distinct requirements). Keep the
    # first row seen per key; due date math (not the row's flag columns)
    # determines status once, after dedup.
    seen: dict[tuple, tuple] = {}
    for row in rows[header_idx + 1:]:
        if not row[c_associate]:  # skip empty rows
            continue
        due_date = row[c_due_date] if isinstance(row[c_due_date], datetime) else None
        key = (row[c_associate], row[c_win], row[c_item_name], due_date)
        if key not in seen:
            seen[key] = row

    records: list[CBLRecord] = []
    today = date.today()
    for row in seen.values():
        due_date = row[c_due_date] if isinstance(row[c_due_date], datetime) else None
        status = _status_from_due_date(due_date, today)
        if status is None:
            continue  # no due date, or beyond the tracked horizon

        manager = str(row[c_manager]).strip() if row[c_manager] else "No Manager"
        shift = _parse_shift(row[c_job_description])

        records.append(CBLRecord(
            associate=str(row[c_associate]).strip(),
            win=row[c_win],
            user_id=str(row[c_user_id]).strip() if row[c_user_id] else "",
            job_description=str(row[c_job_description]).strip() if row[c_job_description] else "",
            item_name=str(row[c_item_name]).strip() if row[c_item_name] else "",
            due_date=due_date,
            status=status,
            manager=manager,
            shift=shift,
        ))

    _cache = records
    _loaded_at = datetime.now()
    return records


def get_summary(records: list[CBLRecord]) -> dict:
    """Build summary stats for the main dashboard."""
    total = len(records)

    # Per manager per status
    manager_stats: dict[str, dict[str, int]] = {}
    for r in records:
        if r.manager not in manager_stats:
            manager_stats[r.manager] = {s: 0 for s in STATUS_ORDER}
        manager_stats[r.manager][r.status] += 1

    # Per shift per status
    shift_stats: dict[str, dict[str, int]] = {}
    for r in records:
        if r.shift not in shift_stats:
            shift_stats[r.shift] = {s: 0 for s in STATUS_ORDER}
        shift_stats[r.shift][r.status] += 1

    # Sort managers by total CBLs desc
    manager_totals = {
        m: sum(v.values()) for m, v in manager_stats.items()
    }
    sorted_managers = sorted(manager_stats.keys(), key=lambda m: manager_totals[m], reverse=True)

    # Sort shifts
    sorted_shifts = sorted(shift_stats.keys())

    return {
        "total": total,
        "manager_stats": manager_stats,
        "manager_totals": manager_totals,
        "sorted_managers": sorted_managers,
        "shift_stats": shift_stats,
        "sorted_shifts": sorted_shifts,
    }


def get_manager_associates(records: list[CBLRecord], manager: str) -> dict:
    """Get all associates + their CBLs for a specific manager."""
    filtered = [r for r in records if r.manager == manager]

    # Group by associate
    assoc_map: dict[str, dict[str, list[CBLRecord]]] = {}
    for r in filtered:
        if r.associate not in assoc_map:
            assoc_map[r.associate] = {s: [] for s in STATUS_ORDER}
        assoc_map[r.associate][r.status].append(r)

    # Sort associates by overdue count desc
    sorted_assocs = sorted(
        assoc_map.keys(),
        key=lambda a: sum(len(v) for v in assoc_map[a].values()),
        reverse=True,
    )

    return {
        "manager": manager,
        "total": len(filtered),
        "assoc_map": assoc_map,
        "sorted_assocs": sorted_assocs,
    }