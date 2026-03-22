"""PHL5 CBL Compliance Dashboard - FastAPI app."""
from urllib.parse import quote, unquote

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from data import get_manager_associates, get_summary, load_data, STATUS_ORDER
from pto_data import (
    get_pto_manager_detail,
    get_pto_summary,
    load_pto,
)
from points_data import (
    get_points_manager_detail,
    get_points_summary,
    load_points,
)
from checkin_data import (
    get_checkin_manager_detail,
    get_checkin_summary,
    load_checkins,
)
from attendance_data import (
    get_attendance_manager_detail,
    get_attendance_summary,
    load_attendance,
    EXCEPTION_LABELS,
)

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import logging

import json
import onedrive_client
from file_watcher import start_file_watcher
from scorecard_data import get_scorecard_summary, get_manager_scorecard
from associate_lookup import search_associate
from history_db import init_db, save_snapshot, get_snapshots
from notes_db import init_notes_table, get_notes_for_wins, save_note
from alerts import check_and_send_alerts
from shifts_data import get_shift_breakdown
from email_scorecards import send_all_scorecards, get_manager_email, _render_scorecard_email, _load_config, send_email

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


from datetime import datetime

_last_refreshed: datetime | None = None


def scheduled_refresh():
    """Auto-refresh all data caches on schedule.

    Downloads fresh bytes from OneDrive once, then re-parses all sheets.
    """
    global _last_refreshed
    logger.info("Scheduled refresh triggered - downloading latest Excel from OneDrive...")
    onedrive_client.refresh_file_bytes()
    load_data(force=True)
    load_attendance(force=True)
    load_checkins(force=True)
    load_points(force=True)
    load_pto(force=True)
    _last_refreshed = datetime.now()
    # Save daily snapshot for trending
    _save_snapshot_now()
    # Check alert thresholds
    try:
        check_and_send_alerts(
            load_data(), load_attendance(), load_checkins(), load_points(), load_pto()
        )
    except Exception as exc:
        logger.warning(f"Alert check failed: {exc}")
    logger.info("Scheduled refresh complete!")


def _save_snapshot_now() -> None:
    """Pull current totals from cached data and persist a snapshot."""
    try:
        cbl = load_data()
        att = load_attendance()
        chk = load_checkins()
        pts = load_points()
        pto = load_pto()
        from checkin_data import get_checkin_summary
        chk_summary = get_checkin_summary(chk)
        save_snapshot(
            cbl=len(cbl),
            att=len(att),
            chk=chk_summary["total"],
            pts=len(pts),
            pto=len(pto),
        )
        logger.info("Snapshot saved to history.db")
    except Exception as exc:
        logger.warning(f"Snapshot failed: {exc}")


app = FastAPI(title="PHL5 Compliance Dashboard")


@app.on_event("startup")
async def startup_event():
    """Authenticate + load all data eagerly on startup.

    This triggers the browser login popup immediately when the server starts
    rather than waiting for the first HTTP request.
    """
    logger.info("Starting up - authenticating with OneDrive...")
    onedrive_client.refresh_file_bytes()   # authenticate + download once
    load_data()
    load_attendance()
    load_checkins()
    load_points()
    load_pto()
    global _last_refreshed
    _last_refreshed = datetime.now()
    init_db()
    init_notes_table()
    _save_snapshot_now()
    logger.info("All data loaded and ready!")
    start_file_watcher(scheduled_refresh)

# Auto-refresh at 8:30 AM and 8:30 PM every day
scheduler = BackgroundScheduler()
scheduler.add_job(scheduled_refresh, CronTrigger(hour=8, minute=30))
scheduler.add_job(scheduled_refresh, CronTrigger(hour=20, minute=30))
scheduler.start()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")
templates.env.globals["quote"] = quote


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    # Load all 3 totals for the home page cards
    cbl_records = load_data()
    cbl_summary = get_summary(cbl_records)

    att_records = load_attendance()
    att_summary = get_attendance_summary(att_records)

    chk_records = load_checkins()
    chk_summary = get_checkin_summary(chk_records)

    pts_records = load_points()
    pts_summary = get_points_summary(pts_records)

    pto_records = load_pto()
    pto_summary = get_pto_summary(pto_records)

    return templates.TemplateResponse("home.html", {
        "request": request,
        "cbl_total": cbl_summary["total"],
        "att_total": att_summary["total"],
        "chk_total": chk_summary["total"],
        "chk_associates": chk_summary["total_associates"],
        "pts_total": pts_summary["total"],
        "pto_total": pto_summary["total"],
    })


@app.get("/cbls", response_class=HTMLResponse)
async def index(request: Request):
    records = load_data()
    summary = get_summary(records)
    return templates.TemplateResponse("index.html", {
        "request": request,
        "summary": summary,
        "status_order": STATUS_ORDER,
        "total": summary["total"],
    })


@app.get("/cbls/manager/{manager_name}", response_class=HTMLResponse)
async def manager_detail(request: Request, manager_name: str):
    manager = unquote(manager_name)
    records = load_data()
    data = get_manager_associates(records, manager)
    return templates.TemplateResponse("manager.html", {
        "request": request,
        "data": data,
        "status_order": STATUS_ORDER,
        "quote": quote,
    })


@app.get("/attendance", response_class=HTMLResponse)
async def attendance(request: Request):
    records = load_attendance()
    summary = get_attendance_summary(records)
    return templates.TemplateResponse("attendance.html", {
        "request": request,
        "summary": summary,
        "labels": EXCEPTION_LABELS,
    })


@app.get("/attendance/manager/{manager_name}", response_class=HTMLResponse)
async def attendance_manager(request: Request, manager_name: str):
    manager = unquote(manager_name)
    records = load_attendance()
    data = get_attendance_manager_detail(records, manager)
    return templates.TemplateResponse("attendance_manager.html", {
        "request": request,
        "data": data,
        "labels": EXCEPTION_LABELS,
        "quote": quote,
    })


@app.get("/checkins", response_class=HTMLResponse)
async def checkins(request: Request):
    records = load_checkins()
    summary = get_checkin_summary(records)
    return templates.TemplateResponse("checkins.html", {
        "request": request,
        "summary": summary,
    })


@app.get("/checkins/manager/{manager_name}", response_class=HTMLResponse)
async def checkins_manager(request: Request, manager_name: str):
    manager = unquote(manager_name)
    records = load_checkins()
    data = get_checkin_manager_detail(records, manager)
    return templates.TemplateResponse("checkins_manager.html", {
        "request": request,
        "data": data,
        "quote": quote,
    })


@app.get("/points", response_class=HTMLResponse)
async def points(request: Request):
    import json
    records = load_points()
    summary = get_points_summary(records)
    records_json = json.dumps([
        {"associate": r.associate, "win": r.win, "manager": r.manager,
         "shift": r.shift, "occurrences": r.occurrences}
        for r in records
    ])
    chart_managers = json.dumps(summary["sorted_managers"])
    chart_occs = json.dumps([
        summary["manager_stats"][m]["max_occ"] for m in summary["sorted_managers"]
    ])
    return templates.TemplateResponse("points.html", {
        "request": request,
        "summary": summary,
        "records_json": records_json,
        "chart_managers": chart_managers,
        "chart_occs": chart_occs,
    })


@app.get("/points/manager/{manager_name}", response_class=HTMLResponse)
async def points_manager(request: Request, manager_name: str):
    manager = unquote(manager_name)
    records = load_points()
    data = get_points_manager_detail(records, manager)
    wins = [str(r.win) for r in data["records"]]
    notes = get_notes_for_wins(wins)
    return templates.TemplateResponse("points_manager.html", {
        "request": request,
        "data": data,
        "notes": notes,
        "quote": quote,
    })


@app.post("/api/notes/{win}")
async def upsert_note(win: str, request: Request):
    form = await request.form()
    note_text = form.get("note", "").strip()
    save_note(win, note_text)
    if note_text:
        return HTMLResponse(
            f'<span class="inline-flex items-center gap-1 bg-blue-50 text-[#0053e2] '
            f'border border-blue-100 rounded-lg px-2 py-1 text-xs">'
            f'📝 {note_text}</span> '
            f'<span class="text-green-600 text-xs font-semibold">✓ Saved</span>'
        )
    return HTMLResponse('<span class="text-gray-400 text-xs">Note cleared</span>')


@app.get("/pto", response_class=HTMLResponse)
async def pto(request: Request):
    import json
    records = load_pto()
    summary = get_pto_summary(records)
    records_json = json.dumps([
        {"date_submitted": r.date_submitted, "date_requested": r.date_requested,
         "associate": r.associate, "win": r.win, "position": r.position,
         "time_requested": r.time_requested, "manager": r.manager, "shift": r.shift}
        for r in records
    ])
    return templates.TemplateResponse("pto.html", {
        "request": request,
        "summary": summary,
        "records_json": records_json,
    })


@app.get("/pto/manager/{manager_name}", response_class=HTMLResponse)
async def pto_manager(request: Request, manager_name: str):
    manager = unquote(manager_name)
    records = load_pto()
    data = get_pto_manager_detail(records, manager)
    return templates.TemplateResponse("pto_manager.html", {
        "request": request,
        "data": data,
        "quote": quote,
    })


@app.get("/shifts", response_class=HTMLResponse)
async def shifts(request: Request):
    cbl = load_data()
    att = load_attendance()
    pts = load_points()
    pto = load_pto()
    breakdown = get_shift_breakdown(cbl, att, pts, pto)
    import json as _json
    return templates.TemplateResponse("shifts.html", {
        "request": request,
        "breakdown": breakdown,
        "breakdown_json": _json.dumps(breakdown["rows"]),
    })


@app.post("/send-scorecard/{manager_name}")
async def send_single_scorecard(manager_name: str, request: Request):
    """Send one manager's scorecard to a user-supplied email address."""
    body = await request.json()
    email = (body.get("email") or "").strip()
    if not email or "@" not in email:
        return JSONResponse({"ok": False, "error": "Invalid email address"}, status_code=400)

    manager = unquote(manager_name)
    cbl = load_data()
    att = load_attendance()
    chk = load_checkins()
    pts = load_points()
    pto = load_pto()
    data    = get_manager_scorecard(manager, cbl, att, chk, pts, pto)
    html    = _render_scorecard_email(manager, data)
    subject = f"PHL5 Compliance Scorecard \u2014 {manager} \u2014 {datetime.now().strftime('%b %d, %Y')}"
    try:
        send_email(email, subject, html)
        logger.info(f"Scorecard emailed for {manager} -> {email}")
        return JSONResponse({"ok": True})
    except Exception as exc:
        logger.warning(f"Scorecard email failed for {manager}: {exc}")
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=500)


@app.post("/send-scorecards", response_class=HTMLResponse)
async def send_scorecards_now(request: Request):
    result = send_all_scorecards(dry_run=False)
    return templates.TemplateResponse("send_scorecards_result.html", {
        "request": request,
        "result": result,
    })


@app.get("/leaderboard", response_class=HTMLResponse)
async def leaderboard(request: Request):
    cbl = load_data()
    att = load_attendance()
    chk = load_checkins()
    pts = load_points()
    pto = load_pto()
    summary = get_scorecard_summary(cbl, att, chk, pts, pto)
    rows = summary["rows"]
    # Sort best (fewest issues) → worst
    ranked = sorted(rows, key=lambda x: x["issue_total"])
    worst = ranked[-1]["issue_total"] if ranked else 1
    if worst == 0:
        worst = 1
    for i, r in enumerate(ranked):
        r["score"] = round((1 - r["issue_total"] / worst) * 100)
        r["rank"] = i + 1
    return templates.TemplateResponse("leaderboard.html", {
        "request": request,
        "rows": ranked,
    })


@app.get("/api/last-refreshed", response_class=HTMLResponse)
async def last_refreshed():
    """Returns a small HTML snippet showing the last refresh time."""
    if _last_refreshed:
        ts = _last_refreshed.strftime("%b %d, %Y %I:%M %p")
        return HTMLResponse(f'<span>&#128260; Last refreshed: <strong>{ts}</strong></span>')
    return HTMLResponse('<span>&#128260; Not yet refreshed</span>')


@app.get("/lookup", response_class=HTMLResponse)
async def lookup(request: Request, q: str = ""):
    cbl = load_data()
    att = load_attendance()
    chk = load_checkins()
    pts = load_points()
    pto = load_pto()
    results = search_associate(q, cbl, att, chk, pts, pto) if q else None
    return templates.TemplateResponse("lookup.html", {
        "request": request,
        "q": q,
        "results": results,
        "quote": quote,
    })


@app.get("/trends", response_class=HTMLResponse)
async def trends(request: Request):
    snapshots = get_snapshots(days=60)
    return templates.TemplateResponse("trends.html", {
        "request": request,
        "snapshots": snapshots,
        "snapshots_json": json.dumps(snapshots),
    })


@app.get("/scorecard", response_class=HTMLResponse)
async def scorecard(request: Request):
    cbl = load_data()
    att = load_attendance()
    chk = load_checkins()
    pts = load_points()
    pto = load_pto()
    summary = get_scorecard_summary(cbl, att, chk, pts, pto)
    return templates.TemplateResponse("scorecard.html", {
        "request": request,
        "summary": summary,
        "quote": quote,
    })


@app.get("/scorecard/manager/{manager_name}", response_class=HTMLResponse)
async def scorecard_manager(request: Request, manager_name: str):
    manager = unquote(manager_name)
    cbl = load_data()
    att = load_attendance()
    chk = load_checkins()
    pts = load_points()
    pto = load_pto()
    data = get_manager_scorecard(manager, cbl, att, chk, pts, pto)
    return templates.TemplateResponse("scorecard_manager.html", {
        "request": request,
        "data": data,
        "quote": quote,
    })


@app.post("/refresh")
async def refresh_all_data():
    """Download latest Excel from OneDrive and reload all data caches."""
    onedrive_client.refresh_file_bytes()
    load_data(force=True)
    load_attendance(force=True)
    load_checkins(force=True)
    load_points(force=True)
    load_pto(force=True)
    return RedirectResponse(url="/", status_code=303)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8501, reload=False)