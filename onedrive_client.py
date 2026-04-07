"""Excel file loader — reads from OneDrive via Microsoft Graph API or local fallback.

Uses the token stored in ~/.code_puppy/msgraph.json (kept fresh by Code Puppy)
to fetch a pre-authenticated Graph download URL for the file each time.

File: PHL5 People Dashboard.xlsx (owned by a0m1czs)
"""
from __future__ import annotations

import io
import json
import logging
from pathlib import Path
from typing import Optional

import openpyxl

logger = logging.getLogger(__name__)

LOCAL_FILE_PATH = Path(__file__).parent / "phl5_compliance.xlsx"
MSGRAPH_TOKEN_FILE = Path.home() / ".code_puppy" / "msgraph.json"

# Stable Graph identifiers for PHL5 People Dashboard.xlsx
_DRIVE_ID = "b!71cpg6Il4U63YQKRP32zw0zAdDbqNsVNuWqDG6i7osRaZ-_vHxs0SJ6Iip-Avq6v"
_ITEM_ID  = "01A5L6XWWRFOGZK3T7KNBLIQ4NWNPOQ4MG"
_GRAPH_ITEM_URL = (
    f"https://graph.microsoft.com/v1.0/drives/{_DRIVE_ID}/items/{_ITEM_ID}"
)

PROXIES = {
    "http": "http://sysproxy.wal-mart.com:8080",
    "https": "http://sysproxy.wal-mart.com:8080",
}

_file_bytes: Optional[bytes] = None


def _get_graph_token() -> str:
    """Read the MS Graph access token stored by Code Puppy."""
    if not MSGRAPH_TOKEN_FILE.exists():
        raise RuntimeError(
            f"MS Graph token not found at {MSGRAPH_TOKEN_FILE}.\n"
            "Open Code Puppy and run any msgraph command to refresh auth."
        )
    data = json.loads(MSGRAPH_TOKEN_FILE.read_text(encoding="utf-8"))
    token = data.get("access_token", "")
    if not token:
        raise RuntimeError("access_token missing in msgraph.json — re-auth via Code Puppy.")
    return token


def _get_graph_download_url(token: str) -> str:
    """Ask Graph API for a fresh pre-authenticated download URL."""
    import requests
    resp = requests.get(
        _GRAPH_ITEM_URL,
        headers={"Authorization": f"Bearer {token}"},
        proxies=PROXIES,
        timeout=15,
    )
    resp.raise_for_status()
    item = resp.json()
    download_url = item.get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise RuntimeError("Graph API response missing @microsoft.graph.downloadUrl")
    return download_url


def _download_from_graph() -> bytes:
    """Fetch fresh bytes from OneDrive via Graph API."""
    import requests
    logger.info("Fetching fresh download URL from Microsoft Graph...")
    token = _get_graph_token()
    download_url = _get_graph_download_url(token)
    logger.info("Downloading Excel from OneDrive via Graph...")
    resp = requests.get(download_url, proxies=PROXIES, timeout=30, allow_redirects=True)
    resp.raise_for_status()
    logger.info(f"Downloaded {len(resp.content):,} bytes from OneDrive")
    return resp.content


def _load_local() -> bytes:
    """Read the local Excel file as bytes."""
    if not LOCAL_FILE_PATH.exists():
        raise FileNotFoundError(
            f"Local Excel file not found: {LOCAL_FILE_PATH}\n"
            "Either place the file there or ensure Code Puppy MS Graph auth is valid."
        )
    logger.info(f"Loading local Excel file: {LOCAL_FILE_PATH}")
    return LOCAL_FILE_PATH.read_bytes()


def refresh_file_bytes() -> None:
    """Reload Excel bytes from OneDrive (Graph API) or fall back to local file."""
    global _file_bytes
    try:
        _file_bytes = _download_from_graph()
        return
    except Exception as exc:
        logger.error(f"OneDrive download failed: {exc}")
        logger.warning("Falling back to local Excel file.")
    _file_bytes = _load_local()


def get_workbook() -> openpyxl.Workbook:
    """Return a fresh openpyxl workbook from cached bytes.

    Loads from OneDrive URL if configured, otherwise reads the local file.
    Each call creates a new Workbook instance so sheets can be fully iterated.
    """
    global _file_bytes
    if _file_bytes is None:
        refresh_file_bytes()
    return openpyxl.load_workbook(io.BytesIO(_file_bytes), data_only=True, read_only=True)
