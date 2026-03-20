"""Generate a polished static HTML snapshot report from the current Excel data.

Usage:
    uv run python export_report.py

Outputs hr_compliance_report.html in the project directory.
"""
from __future__ import annotations

import json
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from data import load_data, get_summary, STATUS_ORDER
from attendance_data import load_attendance, get_attendance_summary, EXCEPTION_LABELS
from checkin_data import load_checkins, get_checkin_summary
from points_data import load_points, get_points_summary
from pto_data import load_pto, get_pto_summary

OUT_PATH = Path(__file__).parent / "hr_compliance_report.html"
GENERATED_AT = datetime.now().strftime("%B %d, %Y at %I:%M %p")

STATUS_COLORS = {
    "Overdue":  "#ea1100",
    "7 Days":   "#c2410c",
    "14 Days":  "#d97706",
    "30 Days":  "#0053e2",
    "60 Days":  "#2a8703",
}


# ---------------------------------------------------------------------------
# HTML primitives
# ---------------------------------------------------------------------------

def _card(label: str, value: int | str, color: str, icon: str = "") -> str:
    return f"""
      <div style="background:{color};color:#fff;border-radius:12px;padding:1.4rem 1.8rem;
                  min-width:150px;flex:1;box-shadow:0 2px 8px rgba(0,0,0,.15);text-align:center">
        <div style="font-size:2.2rem;font-weight:800;letter-spacing:-1px">{value}</div>
        <div style="font-size:0.82rem;margin-top:4px;opacity:.9;font-weight:500">{icon} {label}</div>
      </div>"""


def _cards_row(*cards: str) -> str:
    inner = "".join(cards)
    return f'<div style="display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:1.8rem">{inner}</div>'


def _badge(text: str, color: str) -> str:
    return (f'<span style="display:inline-block;background:{color};color:#fff;'
            f'padding:2px 9px;border-radius:20px;font-size:.75rem;font-weight:700">'
            f'{text}</span>')


def _table(headers: list[str], rows: list[list[str]], max_rows: int = 50) -> str:
    th = "".join(
        f'<th style="background:#0053e2;color:#fff;padding:10px 14px;'
        f'text-align:left;white-space:nowrap">{h}</th>' for h in headers
    )
    tbody = ""
    for i, row in enumerate(rows[:max_rows]):
        bg = "#f0f4ff" if i % 2 == 0 else "#fff"
        tds = "".join(
            f'<td style="padding:8px 14px;border-bottom:1px solid #e5e7eb">{cell}</td>'
            for cell in row
        )
        tbody += f'<tr style="background:{bg}">{tds}</tr>'
    overflow = ""
    if len(rows) > max_rows:
        overflow = (f'<p style="color:#6b7280;font-size:.82rem;margin-top:.5rem">'
                    f'Showing top {max_rows} of {len(rows)} rows.</p>')
    return (
        '<div style="overflow-x:auto">'
        f'<table style="width:100%;border-collapse:collapse;font-size:.88rem;">'
        f'<thead><tr>{th}</tr></thead><tbody>{tbody}</tbody></table></div>{overflow}'
    )


def _section(title: str, body: str, icon: str = "") -> str:
    return f"""
    <section style="background:#fff;border-radius:12px;padding:1.8rem 2rem;
                   box-shadow:0 1px 6px rgba(0,0,0,.08);margin-bottom:2rem">
      <h2 style="margin:0 0 1.2rem;color:#0053e2;font-size:1.2rem;display:flex;
                 align-items:center;gap:.5rem">{icon} {title}</h2>
      {body}
    </section>"""


def _chart_container(chart_id: str, height: str = "280px") -> str:
    return (f'<div style="height:{height};margin-bottom:1.5rem">'
            f'<canvas id="{chart_id}"></canvas></div>')


# ---------------------------------------------------------------------------
# Section builders
# ---------------------------------------------------------------------------

def build_cbl(scripts: list[str]) -> str:
    records = load_data()
    summary = get_summary(records)

    overdue = sum(1 for r in records if r.status == "Overdue")
    due7 = sum(1 for r in records if r.status == "7 Days")

    cards = _cards_row(
        _card("Total CBLs Due", summary["total"], "#0053e2", "📋"),
        _card("Overdue", overdue, "#ea1100", "🔴"),
        _card("Due in 7 Days", due7, "#c2410c", "⚠️"),
        _card("Due in 14 Days", sum(1 for r in records if r.status == "14 Days"), "#d97706", "📅"),
    )

    # Chart data — top 10 managers by total
    top_mgrs = summary["sorted_managers"][:10]
    chart_labels = json.dumps(top_mgrs)
    chart_data = {s: json.dumps([summary["manager_stats"][m][s] for m in top_mgrs]) for s in STATUS_ORDER}

    chart = _chart_container("cblChart", "300px")
    scripts.append(f"""
    new Chart(document.getElementById('cblChart'), {{
      type: 'bar',
      data: {{
        labels: {chart_labels},
        datasets: [
          {{ label:'Overdue', data:{chart_data['Overdue']}, backgroundColor:'{STATUS_COLORS['Overdue']}' }},
          {{ label:'7 Days',  data:{chart_data['7 Days']},  backgroundColor:'{STATUS_COLORS['7 Days']}' }},
          {{ label:'14 Days', data:{chart_data['14 Days']}, backgroundColor:'{STATUS_COLORS['14 Days']}' }},
          {{ label:'30 Days', data:{chart_data['30 Days']}, backgroundColor:'{STATUS_COLORS['30 Days']}' }},
          {{ label:'60 Days', data:{chart_data['60 Days']}, backgroundColor:'{STATUS_COLORS['60 Days']}' }},
        ]
      }},
      options: {{ responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ position:'bottom' }}, title:{{ display:true, text:'CBLs by Manager (Top 10)' }} }},
        scales:{{ x:{{ stacked:true }}, y:{{ stacked:true, beginAtZero:true }} }}
      }}
    }});"""
    )

    rows = []
    for mgr in summary["sorted_managers"]:
        stats = summary["manager_stats"][mgr]
        badges = " ".join(
            _badge(f"{s}: {stats[s]}", STATUS_COLORS[s])
            for s in STATUS_ORDER if stats[s] > 0
        )
        rows.append([mgr, str(summary["manager_totals"][mgr]), badges or "—"])

    tbl = _table(["Manager", "Total CBLs", "Status Breakdown"], rows)
    return _section("CBL Compliance", cards + chart + tbl, "📋")


def build_attendance(scripts: list[str]) -> str:
    records = load_attendance()
    summary = get_attendance_summary(records)

    absent = summary["type_totals"].get("AT_ABSENT", 0)
    ncns = summary["type_totals"].get("AT_ABSENT_NO_CALL", 0)
    late = summary["type_totals"].get("AT_LATE_IN", 0)

    cards = _cards_row(
        _card("Total Exceptions", summary["total"], "#ea1100", "🚨"),
        _card("Absent", absent, "#b91c1c", "❌"),
        _card("No Call No Show", ncns, "#7f1d1d", "📵"),
        _card("Late In", late, "#d97706", "⏰"),
    )

    top_mgrs = summary["sorted_managers"][:10]
    chart_labels = json.dumps(top_mgrs)
    chart_totals = json.dumps([summary["manager_totals"][m] for m in top_mgrs])
    chart = _chart_container("attChart", "260px")
    scripts.append(f"""
    new Chart(document.getElementById('attChart'), {{
      type: 'bar',
      data: {{
        labels: {chart_labels},
        datasets: [{{ label:'Exceptions', data:{chart_totals},
          backgroundColor:'#ea1100', borderRadius:4 }}]
      }},
      options: {{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
        plugins:{{ legend:{{ display:false }}, title:{{ display:true, text:'Top Managers by Attendance Exceptions' }} }},
        scales:{{ x:{{ beginAtZero:true }} }}
      }}
    }});"""
    )

    rows = []
    for mgr in summary["sorted_managers"]:
        stats = summary["manager_stats"][mgr]
        badges = " ".join(
            _badge(f"{EXCEPTION_LABELS.get(t,t)}: {stats[t]}", "#ea1100")
            for t in summary["all_types"] if stats.get(t, 0) > 0
        )
        rows.append([mgr, str(summary["manager_totals"][mgr]), badges or "—"])

    tbl = _table(["Manager", "Total", "Exception Breakdown"], rows)
    return _section("Attendance Exceptions", cards + chart + tbl, "🚨")


def build_checkins(scripts: list[str]) -> str:
    records = load_checkins()
    summary = get_checkin_summary(records)

    cards = _cards_row(
        _card("Associates Needing Check-Ins", summary["total_associates"], "#0053e2", "👥"),
        _card("Total Check-Ins Needed", summary["total"], "#c2410c", "✅"),
    )

    top_mgrs = summary["sorted_managers"][:10]
    chart_labels = json.dumps(top_mgrs)
    chart_totals = json.dumps([summary["manager_totals"][m] for m in top_mgrs])
    chart = _chart_container("chkChart", "260px")
    scripts.append(f"""
    new Chart(document.getElementById('chkChart'), {{
      type: 'bar',
      data: {{
        labels: {chart_labels},
        datasets: [{{ label:'Check-Ins Needed', data:{chart_totals},
          backgroundColor:'#0053e2', borderRadius:4 }}]
      }},
      options: {{ responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }}, title:{{ display:true, text:'Check-Ins Needed by Manager' }} }},
        scales:{{ x:{{ stacked:false }}, y:{{ beginAtZero:true }} }}
      }}
    }});"""
    )

    rows = []
    for mgr in summary["sorted_managers"]:
        needed = summary["manager_totals"][mgr]
        assocs = sum(summary["manager_stats"][mgr].values())
        rows.append([mgr, str(assocs), str(needed)])
    tbl = _table(["Manager", "Associates", "Total Check-Ins Needed"], rows)
    return _section("Manager Check-Ins", cards + chart + tbl, "✅")


def build_points(scripts: list[str]) -> str:
    records = load_points()
    summary = get_points_summary(records)

    cards = _cards_row(
        _card("Associates at 5+ Points", summary["total"], "#ea1100", "⚡"),
    )

    top_mgrs = summary["sorted_managers"][:10]
    chart_labels = json.dumps(top_mgrs)
    chart_totals = json.dumps([summary["manager_stats"][m]["total"] for m in top_mgrs])
    chart = _chart_container("ptsChart", "260px")
    scripts.append(f"""
    new Chart(document.getElementById('ptsChart'), {{
      type: 'bar',
      data: {{
        labels: {chart_labels},
        datasets: [{{ label:'Associates 5+ Points', data:{chart_totals},
          backgroundColor:'#ea1100', borderRadius:4 }}]
      }},
      options: {{ responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ display:false }}, title:{{ display:true, text:'Associates with 5+ Points by Manager' }} }},
        scales:{{ y:{{ beginAtZero:true }} }}
      }}
    }});"""
    )

    rows = []
    for mgr in summary["sorted_managers"]:
        st = summary["manager_stats"][mgr]
        rows.append([mgr, str(st["total"]), f"{st['max_occ']:.1f}"])
    tbl = _table(["Manager", "Associates", "Highest Occurrences"], rows)
    return _section("Attendance Points (5+)", cards + chart + tbl, "⚡")


def build_pto(scripts: list[str]) -> str:
    records = load_pto()
    summary = get_pto_summary(records)

    cards = _cards_row(
        _card("Pending PTO Requests", summary["total"], "#0053e2", "🏖️"),
    )

    top_shifts = sorted(summary["shift_stats"], key=lambda s: summary["shift_stats"][s], reverse=True)[:8]
    pie_labels = json.dumps(top_shifts)
    pie_data = json.dumps([summary["shift_stats"][s] for s in top_shifts])
    pie_colors = json.dumps(["#0053e2","#ffc220","#2a8703","#ea1100","#7c3aed","#0891b2","#d97706","#6b7280"])
    chart = _chart_container("ptoChart", "260px")
    scripts.append(f"""
    new Chart(document.getElementById('ptoChart'), {{
      type: 'doughnut',
      data: {{
        labels: {pie_labels},
        datasets: [{{ data:{pie_data}, backgroundColor:{pie_colors} }}]
      }},
      options: {{ responsive:true, maintainAspectRatio:false,
        plugins:{{ legend:{{ position:'right' }}, title:{{ display:true, text:'PTO Requests by Shift' }} }}
      }}
    }});"""
    )

    rows = [[mgr, str(summary["manager_stats"][mgr])] for mgr in summary["sorted_managers"]]
    tbl = _table(["Manager", "Pending PTO Requests"], rows)
    return _section("Pending Time Off Requests", cards + chart + tbl, "🏖️")


# ---------------------------------------------------------------------------
# Top summary bar
# ---------------------------------------------------------------------------

def build_overview() -> str:
    cbl_total = get_summary(load_data())["total"]
    att_total = get_attendance_summary(load_attendance())["total"]
    chk = get_checkin_summary(load_checkins())
    pts_total = get_points_summary(load_points())["total"]
    pto_total = get_pto_summary(load_pto())["total"]

    return _cards_row(
        _card("CBLs Due", cbl_total, "#0053e2", "📋"),
        _card("Attendance Exceptions", att_total, "#ea1100", "🚨"),
        _card("Check-Ins Needed", chk["total"], "#c2410c", "✅"),
        _card("At 5+ Points", pts_total, "#7f1d1d", "⚡"),
        _card("Pending PTO", pto_total, "#2a8703", "🏖️"),
    )


# ---------------------------------------------------------------------------
# Full page
# ---------------------------------------------------------------------------

def generate() -> str:
    scripts: list[str] = []

    overview = build_overview()
    cbl = build_cbl(scripts)
    att = build_attendance(scripts)
    chk = build_checkins(scripts)
    pts = build_points(scripts)
    pto = build_pto(scripts)

    init_js = "\n".join(scripts)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>PHL5 HR Compliance Report</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
  <style>
    *, *::before, *::after {{ box-sizing: border-box }}
    body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
            background:#f0f4ff; color:#1f2937; font-size:15px }}
    a {{ color:#0053e2 }}
    .topbar {{ background:#0053e2; color:#fff; padding:.75rem 2rem;
               display:flex; align-items:center; justify-content:space-between }}
    .topbar h1 {{ margin:0; font-size:1.25rem; font-weight:700; display:flex; align-items:center; gap:.6rem }}
    .topbar .meta {{ font-size:.8rem; opacity:.8 }}
    .spark-bar {{ background:#ffc220; height:5px }}
    .container {{ max-width:1200px; margin:0 auto; padding:2rem 1.5rem }}
    .nav-pills {{ display:flex; gap:.5rem; flex-wrap:wrap; margin-bottom:2rem }}
    .nav-pills a {{ background:#fff; color:#0053e2; border:2px solid #0053e2;
                    padding:.35rem .9rem; border-radius:20px; text-decoration:none;
                    font-size:.83rem; font-weight:600; transition:all .15s }}
    .nav-pills a:hover {{ background:#0053e2; color:#fff }}
    .overview-label {{ font-size:.95rem; font-weight:700; color:#374151; margin-bottom:.75rem }}
    footer {{ text-align:center; color:#9ca3af; font-size:.78rem; padding:2rem 1rem }}
    @media(max-width:600px) {{ .container {{ padding:1rem }} }}
  </style>
</head>
<body>
<div class="topbar">
  <h1>
    <svg width="28" height="28" viewBox="0 0 100 100" fill="none">
      <circle cx="50" cy="50" r="48" fill="#ffc220"/>
      <text x="50" y="67" text-anchor="middle" font-size="52" font-weight="900" fill="#0053e2">W</text>
    </svg>
    PHL5 HR Compliance Report
  </h1>
  <div class="meta">Generated {GENERATED_AT}</div>
</div>
<div class="spark-bar"></div>

<div class="container">

  <nav class="nav-pills" aria-label="Jump to section">
    <a href="#cbl">📋 CBL Compliance</a>
    <a href="#att">🚨 Attendance Exceptions</a>
    <a href="#chk">✅ Manager Check-Ins</a>
    <a href="#pts">⚡ Attendance Points</a>
    <a href="#pto">🏖️ Pending PTO</a>
  </nav>

  <div class="overview-label">Facility Overview</div>
  {overview}

  <div id="cbl">{cbl}</div>
  <div id="att">{att}</div>
  <div id="chk">{chk}</div>
  <div id="pts">{pts}</div>
  <div id="pto">{pto}</div>

  <footer>
    PHL5 HR Compliance Dashboard &bull; Auto-generated {GENERATED_AT} &bull;
    Data source: PHL5 People Dashboard.xlsx
  </footer>
</div>

<script>
{init_js}
</script>
</body>
</html>"""


if __name__ == "__main__":
    print("Generating report...")
    html = generate()
    OUT_PATH.write_text(html, encoding="utf-8")
    print(f"Saved: {OUT_PATH}")
