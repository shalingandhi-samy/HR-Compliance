"""SQLite persistence for manager notes on associates."""
from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "history.db"


def init_notes_table() -> None:
    """Create the notes table if it doesn't exist."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS associate_notes (
                win       TEXT PRIMARY KEY,
                note      TEXT NOT NULL DEFAULT '',
                updated_at TEXT
            )
        """)


def get_note(win: str) -> str:
    """Return the note for a given associate WIN (empty string if none)."""
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT note FROM associate_notes WHERE win = ?", (win,)
        ).fetchone()
    return row[0] if row else ""


def get_notes_for_wins(wins: list[str]) -> dict[str, str]:
    """Bulk-fetch notes for a list of WINs — returns {win: note}."""
    if not wins:
        return {}
    placeholders = ",".join("?" * len(wins))
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            f"SELECT win, note FROM associate_notes WHERE win IN ({placeholders})",
            wins,
        ).fetchall()
    return {r[0]: r[1] for r in rows}


def save_note(win: str, note: str) -> None:
    """Upsert a note for a given associate WIN."""
    now = datetime.now().isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            INSERT INTO associate_notes (win, note, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(win) DO UPDATE SET
                note=excluded.note,
                updated_at=excluded.updated_at
        """, (win, note.strip(), now))