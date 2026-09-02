"""Excel file loader — reads PHL5 People Dashboard from SharePoint via MS Graph.

Token management:
- Fully self-sufficient. Bootstraps a refresh_token once from Code Puppy's
  ~/.code_puppy/msgraph.json (MSAL cache), then owns its own token lifecycle
  from there on via the OAuth refresh_token grant -- no Code Puppy needed
  for routine operation. Own state persists in .graph_token_cache.json
  (gitignored) so it survives process restarts.
- Falls back to the local phl5_compliance.xlsx if Graph is unreachable and
  no usable refresh token exists anywhere (e.g. it's finally expired/revoked
  after long inactivity -- rare, but if it happens, chat with Code Puppy to
  re-bootstrap via msgraph).

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
OWN_TOKEN_CACHE_FILE = Path(__file__).parent / ".graph_token_cache.json"

# SharePoint: 7381HRClerk — Shared Documents/PHL5 People Dashboard.xlsx
_DRIVE_ID = "b!71cpg6Il4U63YQKRP32zw0zAdDbqNsVNuWqDG6i7osRaZ-_vHxs0SJ6Iip-Avq6v"
_ITEM_ID  = "01A5L6XWWRFOGZK3T7KNBLIQ4NWNPOQ4MG"
_GRAPH_ITEM_URL = (
    f"https://graph.microsoft.com/v1.0/drives/{_DRIVE_ID}/items/{_ITEM_ID}"
)

# Same Azure AD app + tenant Code Puppy's own msgraph tooling uses. We
# bootstrap a refresh_token from its cache once, then own renewal ourselves.
_TENANT_ID = "3cbcc3d3-094d-4006-9849-0d11d61f484d"
_CLIENT_ID = "c9516dcf-d06d-487a-b6c0-6c44e3b52a26"
_TOKEN_URL = f"https://login.microsoftonline.com/{_TENANT_ID}/oauth2/v2.0/token"

PROXIES = {
    "http": "http://sysproxy.wal-mart.com:8080",
    "https": "http://sysproxy.wal-mart.com:8080",
}

_file_bytes: Optional[bytes] = None

# In-memory cache so we don't hit disk/network on every single call within
# the same refresh window.
_mem_access_token: Optional[str] = None
_mem_expires_at: Optional[datetime] = None


# ---------------------------------------------------------------------------
# Token helpers
# ---------------------------------------------------------------------------

def _newest_msal_refresh_token(data: dict) -> Optional[str]:
    """Find the refresh_token matching our client_id in Code Puppy's MSAL cache.

    Used only as a one-time bootstrap seed the first time this app runs (or
    if our own persisted cache is ever wiped). After that, we mint our own
    refresh tokens and never need to touch msgraph.json again.
    """
    for entry in data.get("RefreshToken", {}).values():
        if entry.get("client_id") == _CLIENT_ID:
            return entry.get("secret")
    return None


def _load_own_cache() -> Optional[dict]:
    if not OWN_TOKEN_CACHE_FILE.exists():
        return None
    try:
        return json.loads(OWN_TOKEN_CACHE_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None


def _save_own_cache(access_token: str, refresh_token: str, expires_at: datetime) -> None:
    OWN_TOKEN_CACHE_FILE.write_text(
        json.dumps({
            "access_token": access_token,
            "refresh_token": refresh_token,
            "expires_at": expires_at.isoformat(),
        }),
        encoding="utf-8",
    )


def _get_bootstrap_refresh_token() -> str:
    """Get a refresh_token to start from: our own cache, else Code Puppy's."""
    own = _load_own_cache()
    if own and own.get("refresh_token"):
        return own["refresh_token"]

    if not MSGRAPH_TOKEN_FILE.exists():
        raise RuntimeError(
            f"No refresh token available anywhere (checked {OWN_TOKEN_CACHE_FILE} "
            f"and {MSGRAPH_TOKEN_FILE}). Chat with Code Puppy and run any msgraph "
            "command once to bootstrap authentication."
        )
    data = json.loads(MSGRAPH_TOKEN_FILE.read_text(encoding="utf-8"))
    refresh_token = _newest_msal_refresh_token(data)
    if not refresh_token:
        raise RuntimeError(
            "No usable refresh_token found in msgraph.json. Chat with Code "
            "Puppy and run any msgraph command once to bootstrap authentication."
        )
    logger.info("Bootstrapped refresh token from Code Puppy's msgraph.json cache.")
    return refresh_token


def _refresh_access_token() -> tuple[str, datetime]:
    """Mint a fresh access token via the OAuth refresh_token grant.

    Fully self-contained -- no Code Puppy involvement needed. Azure AD
    typically rotates the refresh_token on each use, so we always persist
    whatever comes back (falling back to the old one if none was issued).
    """
    import requests
    refresh_token = _get_bootstrap_refresh_token()
    resp = requests.post(
        _TOKEN_URL,
        data={
            "client_id": _CLIENT_ID,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        },
        proxies=PROXIES,
        timeout=15,
    )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Refresh token grant failed ({resp.status_code}): {resp.text[:300]}\n"
            "The refresh token may finally be expired/revoked. Chat with Code "
            "Puppy and run any msgraph command once to re-bootstrap."
        )
    payload = resp.json()
    access_token = payload["access_token"]
    new_refresh_token = payload.get("refresh_token", refresh_token)
    expires_at = datetime.now() + timedelta(seconds=int(payload.get("expires_in", 3600)))
    _save_own_cache(access_token, new_refresh_token, expires_at)
    logger.info(f"Refreshed Graph access token via refresh_token grant (valid until {expires_at.isoformat()}).")
    return access_token, expires_at


def _get_graph_token() -> str:
    """Return a valid Graph access token, refreshing automatically as needed.

    Fully self-sufficient: checks in-memory cache first, then the persisted
    own-cache file, and only calls the refresh_token grant when the current
    token is missing or within 5 minutes of expiry. No Code Puppy round trip
    required for routine operation.
    """
    global _mem_access_token, _mem_expires_at

    buffer = timedelta(minutes=5)
    now = datetime.now()

    if _mem_access_token and _mem_expires_at and now < _mem_expires_at - buffer:
        return _mem_access_token

    own = _load_own_cache()
    if own:
        try:
            expires_at = datetime.fromisoformat(own["expires_at"])
            if now < expires_at - buffer:
                _mem_access_token = own["access_token"]
                _mem_expires_at = expires_at
                return _mem_access_token
        except (KeyError, ValueError):
            pass

    access_token, expires_at = _refresh_access_token()
    _mem_access_token = access_token
    _mem_expires_at = expires_at
    return access_token


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
