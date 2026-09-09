"""SQLite persistence for daily compliance snapshots (Historical Trending)."""
from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path

import fiscal_calendar

DB_PATH = Path(__file__).parent / "history.db"


def init_db() -> None:
    """Create the snapshots table if it doesn't exist."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS snapshots (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                snapshot_date TEXT NOT NULL UNIQUE,
                cbl_total     INTEGER DEFAULT 0,
                att_total     INTEGER DEFAULT 0,
                chk_total     INTEGER DEFAULT 0,
                pts_total     INTEGER DEFAULT 0,
                pto_total     INTEGER DEFAULT 0,
                created_at    TEXT
            )
        """)


def save_snapshot(
    cbl: int, att: int, chk: int, pts: int, pto: int
) -> None:
    """Upsert today's snapshot — one record per calendar day."""
    today = date.today().isoformat()
    now = datetime.now().isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            INSERT INTO snapshots
                (snapshot_date, cbl_total, att_total, chk_total, pts_total, pto_total, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(snapshot_date) DO UPDATE SET
                cbl_total=excluded.cbl_total,
                att_total=excluded.att_total,
                chk_total=excluded.chk_total,
                pts_total=excluded.pts_total,
                pto_total=excluded.pto_total,
                created_at=excluded.created_at
        """, (today, cbl, att, chk, pts, pto, now))


def get_snapshots(days: int = 60) -> list[dict]:
    """Return the last N days of snapshots, oldest first."""
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute("""
            SELECT snapshot_date, cbl_total, att_total, chk_total, pts_total, pto_total
            FROM snapshots
            ORDER BY snapshot_date DESC
            LIMIT ?
        """, (days,)).fetchall()
    return [
        {"date": r[0], "cbl": r[1], "att": r[2],
         "chk": r[3], "pts": r[4], "pto": r[5]}
        for r in reversed(rows)
    ]


def get_weekly_snapshots(weeks: int = 26) -> list[dict]:
    """Return compliance totals bucketed into Walmart fiscal weeks (Sat-Fri).

    Each week's value is its *latest* snapshot within that week (an
    end-of-week reading), not a sum -- these are point-in-time compliance
    totals, so summing them would be meaningless. Weeks with no snapshot at
    all are simply absent (no fabricated zero/interpolated rows).

    Returns oldest-first, capped at the most recent `weeks` fiscal weeks.
    """
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute("""
            SELECT snapshot_date, cbl_total, att_total, chk_total, pts_total, pto_total
            FROM snapshots
            ORDER BY snapshot_date ASC
        """).fetchall()

    buckets: dict[date, dict] = {}
    for r in rows:
        snap_date = date.fromisoformat(r[0])
        ws = fiscal_calendar.week_start(snap_date)
        # Rows are processed oldest->newest, so the last write per bucket
        # naturally ends up holding that week's latest snapshot.
        buckets[ws] = {
            "week_start": ws.isoformat(),
            "week_label": fiscal_calendar.week_label(snap_date),
            "range_label": fiscal_calendar.week_range_label(ws),
            "as_of": r[0],
            "cbl": r[1], "att": r[2], "chk": r[3], "pts": r[4], "pto": r[5],
        }

    ordered = [buckets[k] for k in sorted(buckets.keys())]
    return ordered[-weeks:]
