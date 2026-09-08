"""Walmart fiscal calendar helpers.

Walmart's retail fiscal week runs Saturday -> Friday. The fiscal year
starts on the Saturday on/before February 1st (that Saturday is the first
day of fiscal Week 1). This matches the confirmed reference point: today
being in fiscal Week 32 for the year anchored at Sat, Jan 31, 2026.
"""
from __future__ import annotations

from datetime import date, timedelta

_SATURDAY = 5  # date.weekday(): Monday=0 ... Saturday=5, Sunday=6


def _fy_anchor(calendar_year: int) -> date:
    """Saturday on/before Feb 1 of the given calendar year -- fiscal Week 1's start."""
    feb1 = date(calendar_year, 2, 1)
    return feb1 - timedelta(days=(feb1.weekday() - _SATURDAY) % 7)


def week_start(d: date) -> date:
    """Return the Saturday that starts d's Walmart fiscal week."""
    return d - timedelta(days=(d.weekday() - _SATURDAY) % 7)


def fiscal_week(d: date) -> tuple[int, int]:
    """Return (fiscal_year, week_number) for date d.

    fiscal_year is named after the calendar year its Feb-1 anchor falls in.
    A fiscal year can run 52 or 53 weeks; 53 is used as a safe upper bound
    when deciding which of this year's / last year's anchor a date belongs to.
    """
    ws = week_start(d)
    for fy in (d.year, d.year - 1):
        anchor = _fy_anchor(fy)
        if anchor <= ws < anchor + timedelta(weeks=53):
            return fy, (ws - anchor).days // 7 + 1
    # Practically unreachable given the +/-1 year search window above.
    anchor = _fy_anchor(d.year)
    return d.year, (ws - anchor).days // 7 + 1


def week_label(d: date) -> str:
    """Human label like 'FY26 Wk32'."""
    fy, wk = fiscal_week(d)
    return f"FY{fy % 100} Wk{wk}"


def week_range_label(ws: date) -> str:
    """Date-range label like 'Sep 05-11' for a week starting on ws (a Saturday)."""
    we = ws + timedelta(days=6)
    if ws.month == we.month:
        return f"{ws.strftime('%b %d')}-{we.strftime('%d')}"
    return f"{ws.strftime('%b %d')}-{we.strftime('%b %d')}"
