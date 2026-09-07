"""Shared helper for turning fragile Excel column *positions* into
resilient column *names*.

The source workbook (PHL5 People Dashboard.xlsx) is maintained by a
different team upstream, and its columns get renamed or reordered without
warning. That's already bitten this dashboard multiple times: a header
renamed "Associates" -> "Associate", a "Date Hired" column inserted into
Manager Check-Ins that silently shifted every column after it, and the
ULearns shift source moving from the Position column to Job Description.

Reading columns by hardcoded numeric index is a landmine that goes off
*silently* -- the loader just returns wrong or empty data with no error,
and nobody notices until someone asks "why does this say 0?".

The fix: locate the header row, build a {header_text: column_index} map
from whatever the header row *actually* says this week, and look columns
up by name (trying a short list of acceptable aliases for known past
renames). If a required column can't be found under any alias, raise
loudly instead of quietly returning 0/empty -- a startup crash you notice
beats a wrong dashboard number nobody notices.
"""
from __future__ import annotations

from typing import Iterable, Optional


def _normalize(text: object) -> str:
    """Lowercase + trim, so trailing spaces / arrows in header cells

    (e.g. "Late ", "Next 14 Days ", "FORMULA Do Not Modify\u2193") don't
    break exact-match comparisons.
    """
    return str(text).strip().lower() if text is not None else ""


def find_header_row(
    rows: list[tuple],
    first_col_aliases: Iterable[str],
    max_scan: int = 10,
) -> tuple[int, dict[str, int]]:
    """Scan the first `max_scan` rows for the header row.

    A row counts as the header row if its first cell matches (or starts
    with) one of `first_col_aliases`, case-insensitively. Returns
    (row_index, header_map) where header_map maps normalized header text
    -> column index, built from that row's *actual* values -- not from
    any assumption about what should be there.
    """
    aliases = [_normalize(a) for a in first_col_aliases]
    for row_idx, row in enumerate(rows[:max_scan]):
        first = _normalize(row[0] if row else None)
        if any(first == a or (a and first.startswith(a)) for a in aliases):
            header_map = {_normalize(cell): idx for idx, cell in enumerate(row) if cell is not None}
            return row_idx, header_map
    raise RuntimeError(
        f"Could not find a header row (looked for first column matching one of "
        f"{list(first_col_aliases)} in the first {max_scan} rows). "
        "The sheet layout may have changed -- check the source Excel."
    )


def col(header_map: dict[str, int], *aliases: str, required: bool = True) -> Optional[int]:
    """Look up a column index by trying each alias, case-insensitively.

    Tries exact match first, then falls back to prefix match (handles
    trailing notes/arrows on header cells like "FORMULA Do Not Modify\u2193").
    Raises RuntimeError if required and nothing matches -- fail loud
    rather than silently defaulting to the wrong column.
    """
    for alias in aliases:
        norm = _normalize(alias)
        if norm in header_map:
            return header_map[norm]
        for header_text, idx in header_map.items():
            if header_text.startswith(norm):
                return idx
    if required:
        raise RuntimeError(
            f"None of the expected column names {aliases} were found in the header row. "
            f"Available columns: {sorted(header_map.keys())}. "
            "The sheet layout may have changed -- check the source Excel."
        )
    return None
