"""Excel file loader — reads PHL5 People Dashboard from SharePoint via MS Graph.

Token management:
- Reads access_token from ~/.code_puppy/msgraph.json (written by Code Puppy).
- If the token is expired, automatically refreshes it using the stored
  refresh_token and writes the new tokens back to msgraph.json.
- Falls back to the local phl5_compliance.xlsx if Graph is unreachable.

SharePoint source: teams.wal-mart.com/sites/7381HRClerk
File: PHL5 People Dashboard.xlsx
"""
from __future__ import annotations

import io
import json
import logging
from datetime import datetime, timedelta
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

# OAuth app (Code Puppy / Graph Explorer — public client, no secret needed)
_CLIENT_ID = "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
_TENANT_ID = "3cbcc3d3-094d-4006-9849-0d11d61f484d"
_TOKEN_URL = f"https://login.microsoftonline.com/{_TENANT_ID}/oauth2/v2.0/token"
_SCOPES    = "https://graph.microsoft.com/.default offline_access"

PROXIES = {
    "http": "http://sysproxy.wal-mart.com:8080",
    "https": "http://sysproxy.wal-mart.com:8080",
}

_file_bytes: Optional[bytes] = None


# ---------------------------------------------------------------------------
# Token helpers
# ---------------------------------------------------------------------------

def _read_token_file() -> dict:
    if not MSGRAPH_TOKEN_FILE.exists():
        raise RuntimeError(
            f"MS Graph token not at {MSGRAPH_TOKEN_FILE}.\n"
            "Open Code Puppy and run any msgraph command to re-authenticate."
        )
    return json.loads(MSGRAPH_TOKEN_FILE.read_text(encoding="utf-8"))


def _is_token_expired(data: dict) -> bool:
    """Return True if the access token has expired (with 60s buffer)."""
    timestamp = data.get("timestamp")
    expires_in = data.get("expires_in", 0)
    if not timestamp:
        return True
    try:
        issued_at = datetime.fromisoformat(timestamp)
        expires_at = issued_at + timedelta(seconds=int(expires_in) - 60)
        return datetime.now() >= expires_at
    except Exception:
        return True


def _refresh_access_token(data: dict) -> dict:
    """Use the refresh_token to obtain a new access_token from Azure AD."""
    import requests

    refresh_token = data.get("refresh_token", "")
    if not refresh_token:
        raise RuntimeError("No refresh_token in msgraph.json — open Code Puppy to re-authenticate.")

    logger.info("Access token expired — refreshing via refresh_token...")
    resp = requests.post(
        _TOKEN_URL,
        data={
            "grant_type": "refresh_token",
            "client_id": _CLIENT_ID,
            "refresh_token": refresh_token,
            "scope": _SCOPES,
        },
        proxies=PROXIES,
        timeout=15,
    )
    resp.raise_for_status()
    new_tokens = resp.json()

    # Merge new tokens back into the existing data and persist
    data["access_token"] = new_tokens["access_token"]
    data["expires_in"]   = new_tokens.get("expires_in", 3600)
    data["timestamp"]    = datetime.now().isoformat()
    if "refresh_token" in new_tokens:
        data["refresh_token"] = new_tokens["refresh_token"]  # rotation

    MSGRAPH_TOKEN_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    logger.info("Token refreshed and saved to msgraph.json.")
    return data


def _get_graph_token() -> str:
    """Return a valid access token, auto-refreshing if expired."""
    data = _read_token_file()
    if _is_token_expired(data):
        data = _refresh_access_token(data)
    token = data.get("access_token", "")
    if not token:
        raise RuntimeError("access_token missing in msgraph.json — re-auth via Code Puppy.")
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
