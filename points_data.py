"""Data loading and processing for Attendance Points 5+ tab."""
from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

import openpyxl

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "phl5_compliance.xlsx")

# Occurrence risk buckets
BUCKETS = [
    (5, 6,  "At Risk",          "#ffc220", "#995213"),
    (7, 9,  "High Risk",        "#f59e0b", "#92400e"),
    (10, 12, "Critical",        "#ea1100", "#7f1d1d"),
    (13, 99, "Termination Risk", "#7f1d1d", "#ffffff"),
]


def get_bucket(occurrences: float) -> dict:
    for lo, hi, label, bg, text in BUCKETS:
        if lo <= occurrences <= hi:
            return {"label": label, "bg": bg, "text": text}
    return {"label": "At Risk", "bg": "#ffc220", "text": "#995213"}


@dataclass
class PointsRecord:
    win: str
    associate: str
    code: str
    occurrences: float
    manager: str
    shift: str
    bucket: str  # label


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
            # Header row has 'Associate WIN' in col 0
            if row[0] == "Associate WIN":
                header_passed = True
            continue

        if not row[0]:  # skip empty rows
            continue

        try:
            occ = float(row[3]) if row[3] is not None else 0.0
        except (ValueError, TypeError):
            occ = 0.0

        # Strip " (United States of America)" from shift
        shift_raw = str(row[5]).strip() if row[5] else "Unknown"
        shift = shift_raw.split(" (")[0].strip()

        b = get_bucket(occ)
        records.append(PointsRecord(
            win=str(row[0]).strip(),
            associate=str(row[1]).strip() if row[1] else "",
            code=str(row[2]).strip() if row[2] else "",
            occurrences=occ,
            manager=str(row[4]).strip() if row[4] else "No Manager",
            shift=shift,
            bucket=b["label"],
        ))

    wb.close()
    # Sort by occurrences desc
    _cache = sorted(records, key=lambda r: r.occurrences, reverse=True)
    _loaded_at = datetime.now()
    return _cache


def get_points_summary(records: list[PointsRecord]) -> dict:
    """Build summary stats for the points dashboard."""
    total = len(records)

    # Bucket counts
    bucket_counts: dict[str, int] = {}
    for _, _, label, bg, text in BUCKETS:
        bucket_counts[label] = 0
    for r in records:
        if r.bucket in bucket_counts:
            bucket_counts[r.bucket] += 1

    # Per manager stats
    manager_stats: dict[str, dict] = {}
    for r in records:
        if r.manager not in manager_stats:
            manager_stats[r.manager] = {
                "total": 0,
                "buckets": {lbl: 0 for _, _, lbl, _, _ in BUCKETS},
                "max_occ": 0.0,
            }
        manager_stats[r.manager]["total"] += 1
        manager_stats[r.manager]["buckets"][r.bucket] = (
            manager_stats[r.manager]["buckets"].get(r.bucket, 0) + 1
        )
        if r.occurrences > manager_stats[r.manager]["max_occ"]:
            manager_stats[r.manager]["max_occ"] = r.occurrences

    sorted_managers = sorted(
        manager_stats.keys(),
        key=lambda m: manager_stats[m]["total"],
        reverse=True,
    )

    # Per shift stats
    shift_stats: dict[str, int] = {}
    for r in records:
        shift_stats[r.shift] = shift_stats.get(r.shift, 0) + 1

    bucket_meta = [
        {"label": lbl, "bg": bg, "text": text}
        for _, _, lbl, bg, text in BUCKETS
    ]

    return {
        "total": total,
        "bucket_counts": bucket_counts,
        "bucket_meta": bucket_meta,
        "manager_stats": manager_stats,
        "sorted_managers": sorted_managers,
        "shift_stats": shift_stats,
    }


def get_points_manager_detail(records: list[PointsRecord], manager: str) -> dict:
    """Get all associates under a manager."""
    filtered = [r for r in records if r.manager == manager]
    sorted_assocs = sorted(filtered, key=lambda r: r.occurrences, reverse=True)

    bucket_counts: dict[str, int] = {lbl: 0 for _, _, lbl, _, _ in BUCKETS}
    for r in filtered:
        if r.bucket in bucket_counts:
            bucket_counts[r.bucket] += 1

    bucket_meta = [
        {"label": lbl, "bg": bg, "text": text}
        for _, _, lbl, bg, text in BUCKETS
    ]

    return {
        "manager": manager,
        "total": len(filtered),
        "records": sorted_assocs,
        "bucket_counts": bucket_counts,
        "bucket_meta": bucket_meta,
    }