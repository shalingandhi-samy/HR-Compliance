"""Data loading and processing for PHL5 Compliance Dashboard."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import openpyxl

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "phl5_compliance.xlsx")

STATUS_COLS = {
    "Overdue": 6,
    "7 Days": 7,
    "14 Days": 8,
    "30 Days": 9,
    "60 Days": 10,
}

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


def _parse_shift(raw: Optional[str]) -> str:
    if not raw:
        return "Unknown"
    # Extract short shift label e.g. "4 - Weekend (United States of America)" -> "Shift 4 - Weekend"
    parts = str(raw).split(" - ", 1)
    if len(parts) == 2:
        num = parts[0].strip()
        name = parts[1].split(" (")[0].strip()
        return f"Shift {num} - {name}"
    return str(raw)


def _determine_status(row_values: tuple) -> Optional[str]:
    for status, col_idx in STATUS_COLS.items():
        val = row_values[col_idx]
        if val is not None and str(val).strip().lower() in ("x", "yes", "true", "1"):
            return status
    return None


def load_data(force: bool = False) -> list[CBLRecord]:
    """Load Excel data, caching in memory."""
    global _cache, _loaded_at
    if _cache and not force:
        return _cache

    records: list[CBLRecord] = []
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True, read_only=True)
    ws = wb.active

    header_found = False
    for row in ws.iter_rows(values_only=True):
        # Find the actual data header row
        if not header_found:
            if row[0] == "Associate":
                header_found = True
            continue

        if not row[0]:  # skip empty rows
            continue

        status = _determine_status(row)
        if status is None:
            continue  # skip rows with no status flag

        manager = str(row[11]).strip() if row[11] else "No Manager"
        shift = _parse_shift(row[12])

        records.append(CBLRecord(
            associate=str(row[0]).strip(),
            win=row[1],
            user_id=str(row[2]).strip() if row[2] else "",
            job_description=str(row[3]).strip() if row[3] else "",
            item_name=str(row[4]).strip() if row[4] else "",
            due_date=row[5] if isinstance(row[5], datetime) else None,
            status=status,
            manager=manager,
            shift=shift,
        ))

    wb.close()
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