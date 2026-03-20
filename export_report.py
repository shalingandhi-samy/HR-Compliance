"""Generate a static HTML snapshot report from the current Excel data.

Usage:
    uv run python export_report.py

Outputs hr_compliance_report.html in the project directory.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from data import load_data, get_summary, STATUS_ORDER
from attendance_data import load_attendance, get_attendance_summary, EXCEPTION_LABELS
from checkin_data import load_checkins, get_checkin_summary
from points_data import load_points, get_points_summary
from pto_data import load_pto, get_pto_summary

OUT_PATH = Path(__file__).parent / "hr_compliance_report.html"
GENERATED_AT = datetime.now().strftime("%B %d, %Y at %I:%M %p")


def _status_color(status: str) -> str:
    return {
        "Overdue": "#ea1100",
        "7 Days": "#b45309",
        "14 Days": "#d97706",
        "30 Days": "#0053e2",
        "60 Days": "#2a8703",
    }.get(status, "#555")


def _badge(value: int | str, color: str) -> str:
    return (
        f'<span style="background:{color};color:#fff;padding:2px 8px;'
        f'border-radius:12px;font-size:0.8rem;font-weight:700">{value}</span>'
    )


def _section(title: str, body: str) -> str:
    return f"""
    <section style="margin-bottom:2.5rem">
      <h2 style="color:#0053e2;border-bottom:2px solid #0053e2;padding-bottom:0.4rem">{title}</h2>
      {body}
    </section>"""


def _summary_cards(*cards: tuple[str, str, str]) -> str:
    """cards = list of (label, value, color)."""
    items = "".join(
        f'<div style="background:{c};color:#fff;border-radius:10px;padding:1.2rem 1.5rem;'
        f'min-width:140px;text-align:center">'
        f'<div style="font-size:2rem;font-weight:800">{v}</div>'
        f'<div style="font-size:0.85rem;margin-top:0.2rem">{l}</div></div>'
        for l, v, c in cards
    )
    return f'<div style="display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:1.5rem">{items}</div>'


def _table(headers: list[str], rows: list[list[str]]) -> str:
    th = "".join(
        f'<th style="background:#0053e2;color:#fff;padding:8px 12px;text-align:left">{h}</th>'
        for h in headers
    )
    tr_rows = ""
    for i, row in enumerate(rows):
        bg = "#f5f7ff" if i % 2 == 0 else "#fff"
        tds = "".join(
            f'<td style="padding:7px 12px;border-bottom:1px solid #e5e7eb">{cell}</td>'
            for cell in row
        )
        tr_rows += f'<tr style="background:{bg}">{tds}</tr>'
    return (
        '<div style="overflow-x:auto">'
        f'<table style="width:100%;border-collapse:collapse;font-size:0.9rem">'
        f"<thead><tr>{th}</tr></thead><tbody>{tr_rows}</tbody></table></div>"
    )


def build_cbl_section() -> str:
    records = load_data()
    summary = get_summary(records)
    cards = _summary_cards(
        ("Total CBLs Due", summary["total"], "#0053e2"),
        ("Overdue", sum(1 for r in records if r.status == "Overdue"), "#ea1100"),
        ("Due in 7 Days", sum(1 for r in records if r.status == "7 Days"), "#b45309"),
    )
    rows = []
    for mgr in summary["sorted_managers"]:
        stats = summary["manager_stats"][mgr]
        badges = " ".join(
            _badge(stats[s], _status_color(s)) for s in STATUS_ORDER if stats[s] > 0
        )
        rows.append([mgr, str(summary["manager_totals"][mgr]), badges])
    tbl = _table(["Manager", "Total CBLs", "Breakdown"], rows)
    return _section("CBL Compliance", cards + tbl)


def build_attendance_section() -> str:
    records = load_attendance()
    summary = get_attendance_summary(records)
    cards = _summary_cards(
        ("Total Exceptions", summary["total"], "#ea1100"),
    )
    rows = []
    for mgr in summary["sorted_managers"]:
        stats = summary["manager_stats"][mgr]
        total = summary["manager_totals"][mgr]
        breakdown = ", ".join(
            f"{EXCEPTION_LABELS.get(t, t)}: {stats[t]}"
            for t in summary["all_types"] if stats.get(t, 0) > 0
        )
        rows.append([mgr, str(total), breakdown or "—"])
    tbl = _table(["Manager", "Total", "Exception Types"], rows)
    return _section("Attendance Exceptions", cards + tbl)


def build_checkins_section() -> str:
    records = load_checkins()
    summary = get_checkin_summary(records)
    cards = _summary_cards(
        ("Associates Needing Check-Ins", summary["total_associates"], "#0053e2"),
        ("Total Check-Ins Needed", summary["total"], "#b45309"),
    )
    rows = []
    for mgr in summary["sorted_managers"]:
        needed = summary["manager_totals"][mgr]
        assoc_count = sum(summary["manager_stats"][mgr].values())
        rows.append([mgr, str(assoc_count), str(needed)])
    tbl = _table(["Manager", "Associates", "Check-Ins Needed"], rows)
    return _section("Manager Check-Ins", cards + tbl)


def build_points_section() -> str:
    records = load_points()
    summary = get_points_summary(records)
    cards = _summary_cards(
        ("Associates at 5+ Points", summary["total"], "#ea1100"),
    )
    rows = []
    for mgr in summary["sorted_managers"]:
        stats = summary["manager_stats"][mgr]
        rows.append([mgr, str(stats["total"]), str(stats["max_occ"])])
    tbl = _table(["Manager", "Associates", "Max Occurrences"], rows)
    return _section("Attendance Points (5+)", cards + tbl)


def build_pto_section() -> str:
    records = load_pto()
    summary = get_pto_summary(records)
    cards = _summary_cards(
        ("Pending PTO Requests", summary["total"], "#0053e2"),
    )
    rows = []
    for mgr in summary["sorted_managers"]:
        rows.append([mgr, str(summary["manager_stats"][mgr])])
    tbl = _table(["Manager", "Pending PTO"], rows)
    return _section("Pending Time Off Requests", cards + tbl)


def generate() -> str:
    cbl = build_cbl_section()
    att = build_attendance_section()
    chk = build_checkins_section()
    pts = build_points_section()
    pto = build_pto_section()
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PHL5 HR Compliance Report</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
           background:#f3f4f6; margin:0; padding:0; color:#1f2937 }}
    .container {{ max-width:1100px; margin:0 auto; padding:2rem 1.5rem }}
    header {{ background:#0053e2; color:#fff; padding:1.5rem 2rem;
              border-radius:12px; margin-bottom:2rem }}
    header h1 {{ margin:0; font-size:1.6rem }}
    header p {{ margin:0.3rem 0 0; opacity:0.85; font-size:0.9rem }}
    section {{ background:#fff; border-radius:10px; padding:1.5rem 2rem;
               box-shadow:0 1px 4px rgba(0,0,0,0.08); margin-bottom:2rem }}
  </style>
</head>
<body>
<div class="container">
  <header>
    <h1>PHL5 HR Compliance Report</h1>
    <p>Generated {GENERATED_AT} &mdash; Data from PHL5 People Dashboard</p>
  </header>
  {cbl}{att}{chk}{pts}{pto}
  <footer style="text-align:center;color:#9ca3af;font-size:0.8rem;padding:1rem">
    Generated by HR Compliance Dashboard &bull; {GENERATED_AT}
  </footer>
</div>
</body>
</html>"""


if __name__ == "__main__":
    print("Generating report...")
    html = generate()
    OUT_PATH.write_text(html, encoding="utf-8")
    print(f"Report saved to: {OUT_PATH}")
