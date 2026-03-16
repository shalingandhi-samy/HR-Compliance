"""PHL5 CBL Compliance Dashboard - FastAPI app."""
from urllib.parse import quote, unquote

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from data import get_manager_associates, get_summary, load_data, STATUS_ORDER
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

app = FastAPI(title="PHL5 Compliance Dashboard")
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    records = load_data()
    summary = get_summary(records)
    return templates.TemplateResponse("index.html", {
        "request": request,
        "summary": summary,
        "status_order": STATUS_ORDER,
        "total": summary["total"],
    })


@app.get("/manager/{manager_name}", response_class=HTMLResponse)
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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8501, reload=False)