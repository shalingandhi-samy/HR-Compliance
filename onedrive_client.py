"""Excel file loader — reads PHL5 People Dashboard from SharePoint via MS Graph.

Token management:
- Reads access_token from ~/.code_puppy/msgraph.json (written by Code Puppy).
- Does NOT attempt DIY token refresh — Code Puppy manages token lifecycle.
  If the token is expired, just chat with Code Puppy to re-authenticate.
- Falls back to the local phl5_compliance.xlsx if Graph is unreachable.

SharePoint source: teams.wal-mart.com/sites/7381HRClerk
File: PHL5 People Dashboard.xlsx
"""
from __future__ import annotations

import io
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl

logger = logging.getLogger(__name__)

LOCAL_FILE_PATH = Path(__file__).parent / "phl5_compliance.xlsx"
MSGRAPH_TOKEN_FILE = Path.home() / ".code_puppy" / "msgraph.json"

# SharePoint: 7381HRClerk — Shared Documents/PHL5 People Dashboard.xlsx
_DRIVE_ID = "b!71cpg6Il4U63YQKRP32zw0zAdDbqNsVNuWqDG6i7osRaZ-_vHxs0SJ6Iip-Avq6v"
_ITEM_ID  = "01A5L6XWWRFOGZK3T7KNBLIQ4NWNPOQ4MG"
_GRAPH_ITEM_URL = (
    f"https://graph.microsoft.com/v1.0/drives/{_DRIVE_ID}/items/{_ITEM_ID}"
)

# Token is managed by Code Puppy — do not attempt independent refresh.

PROXIES = {
    "http": "http://sysproxy.wal-mart.com:8080",
    "https": "http://sysproxy.wal-mart.com:8080",
}

_file_bytes: Optional[bytes] = None


# ---------------------------------------------------------------------------
# Token helpers
# ---------------------------------------------------------------------------

def _get_graph_token() -> str:
    """Return the current access token from Code Puppy's msgraph.json.

    Token refresh is handled exclusively by Code Puppy — this function
    simply reads what's there. If the token is expired, the Graph API
    call will return 401 and we fall back to the local Excel file.
    To refresh, just chat with Code Puppy (any msgraph command works).
    """
    if not MSGRAPH_TOKEN_FILE.exists():
        raise RuntimeError(
            f"MS Graph token not found at {MSGRAPH_TOKEN_FILE}.\n"
            "Chat with Code Puppy and run any msgraph command to authenticate."
        )
    data = json.loads(MSGRAPH_TOKEN_FILE.read_text(encoding="utf-8"))
    token = data.get("access_token", "")
    if not token:
        raise RuntimeError("access_token missing in msgraph.json — chat with Code Puppy to re-auth.")
    logger.info(f"Using token from msgraph.json (issued {data.get('timestamp', 'unknown')})")
    return token


# ---------------------------------------------------------------------------
# Download helpers
# ---------------------------------------------------------------------------

def _get_graph_download_url(token: str) -> str:
    """Ask Graph API for a pre-authenticated download URL for the Excel file."""
    import requests
    resp = requests.get(
        _GRAPH_ITEM_URL,
        headers={"Authorization": f"Bearer {token}"},
        proxies=PROXIES,
        timeout=15,
    )
    resp.raise_for_status()
    download_url = resp.json().get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise RuntimeError("Graph API response missing @microsoft.graph.downloadUrl")
    return download_url


def _download_from_graph() -> bytes:
    """Fetch fresh Excel bytes from SharePoint via Graph API."""
    import requests
    logger.info("Fetching fresh download URL from Microsoft Graph...")
    token = _get_graph_token()
    download_url = _get_graph_download_url(token)
    logger.info("Downloading Excel from SharePoint via Graph...")
    resp = requests.get(download_url, proxies=PROXIES, timeout=60, allow_redirects=True)
    resp.raise_for_status()
    logger.info(f"Downloaded {len(resp.content):,} bytes from SharePoint")
    return resp.content


def _load_local() -> bytes:
    """Read the local fallback Excel file as bytes."""
    if not LOCAL_FILE_PATH.exists():
        raise FileNotFoundError(
            f"Local Excel file not found: {LOCAL_FILE_PATH}\n"
            "Ensure Code Puppy MS Graph auth is valid or place the file manually."
        )
    logger.info(f"Loading local Excel file: {LOCAL_FILE_PATH}")
    return LOCAL_FILE_PATH.read_bytes()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def refresh_file_bytes() -> None:
    """Reload Excel bytes from SharePoint (auto-refreshing token) or local fallback."""
    global _file_bytes
    try:
        _file_bytes = _download_from_graph()
        return
    except Exception as exc:
        logger.error(f"SharePoint download failed: {exc}")
        logger.warning("Falling back to local Excel file.")
    _file_bytes = _load_local()


def get_workbook() -> openpyxl.Workbook:
    """Return a fresh openpyxl Workbook from cached bytes.

    Each call creates a new instance so sheets can be fully iterated.
    """
    global _file_bytes
    if _file_bytes is None:
        refresh_file_bytes()
    return openpyxl.load_workbook(io.BytesIO(_file_bytes), data_only=True, read_only=True)
