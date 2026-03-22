"""SQLite persistence for daily compliance snapshots (Historical Trending)."""
from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path

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
