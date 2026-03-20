"""PHL5 CBL Compliance Dashboard - FastAPI app."""
from urllib.parse import quote, unquote

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, RedirectResponse
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

import onedrive_client

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def scheduled_refresh():
    """Auto-refresh all data caches on schedule.

    Downloads fresh bytes from OneDrive once, then re-parses all sheets.
    """
    logger.info("⏰ Scheduled refresh triggered — downloading latest Excel from OneDrive...")
    onedrive_client.refresh_file_bytes()
    load_data(force=True)
    load_attendance(force=True)
    load_checkins(force=True)
    load_points(force=True)
    load_pto(force=True)
    logger.info("✅ Scheduled refresh complete!")


app = FastAPI(title="PHL5 Compliance Dashboard")

# Auto-refresh at 8:30 AM and 8:30 PM every day
scheduler = BackgroundScheduler()
scheduler.add_job(scheduled_refresh, CronTrigger(hour=8, minute=30))
scheduler.add_job(scheduled_refresh, CronTrigger(hour=20, minute=30))
scheduler.start()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


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
    return templates.TemplateResponse("points_manager.html", {
        "request": request,
        "data": data,
        "quote": quote,
    })


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