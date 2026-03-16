"""PHL5 CBL Compliance Dashboard - FastAPI app."""
from urllib.parse import quote, unquote

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from data import get_manager_associates, get_summary, load_data, STATUS_ORDER

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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8501, reload=False)