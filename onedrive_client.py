"""Microsoft OneDrive file client using MSAL device code flow.

Downloads the live Excel from OneDrive and caches raw bytes in memory.
All data modules call get_workbook() instead of opening a local file.
"""
from __future__ import annotations

import base64
import io
import logging
import os
from pathlib import Path
from typing import Optional

import msal
import openpyxl
import requests

logger = logging.getLogger(__name__)

# ── Config (override via env vars if needed) ──────────────────────────────────
TENANT_ID = os.getenv("AZURE_TENANT_ID", "wmsc.wal-mart.com")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "d3590ed6-52b3-4102-aeff-aad2292ab01c")  # MS Office public client
SCOPES = ["Files.Read", "Files.Read.All"]

# File owner derived from the OneDrive URL path (a0m1czs_wmsc_wal-mart_com)
FILE_OWNER_EMAIL = os.getenv("ONEDRIVE_FILE_OWNER", "a0m1czs@wmsc.wal-mart.com")
FILE_PATH_IN_DRIVE = "Documents/PHL5 People Dashboard.xlsx"

# Sharing URL (from the link shared by the user)
FILE_SHARING_URL = (
    "https://my.wal-mart.com/:x:/r/personal/a0m1czs_wmsc_wal-mart_com"
    "/Documents/PHL5%20People%20Dashboard.xlsx"
    "?d=wb48b660c85bd404a8a6fa4d17c2be4d4&csf=1&web=1&e=KBu1fF"
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE_PATH = Path(__file__).parent / ".token_cache.json"
LOCAL_FALLBACK_PATH = Path(__file__).parent / "phl5_people_dashboard.xlsx"

PROXIES = {
    "http": "http://sysproxy.wal-mart.com:8080",
    "https": "http://sysproxy.wal-mart.com:8080",
}

# ── Module-level state ────────────────────────────────────────────────────────
_file_bytes: Optional[bytes] = None
_msal_app: Optional[msal.PublicClientApplication] = None
_token_cache = msal.SerializableTokenCache()


# ── MSAL helpers ──────────────────────────────────────────────────────────────

def _get_msal_app() -> msal.PublicClientApplication:
    global _msal_app, _token_cache
    if _msal_app is None:
        if TOKEN_CACHE_PATH.exists():
            _token_cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
        _msal_app = msal.PublicClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            token_cache=_token_cache,
        )
    return _msal_app


def _save_token_cache() -> None:
    if _token_cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(_token_cache.serialize(), encoding="utf-8")
        logger.info("🔑 Token cache saved to disk.")


def _get_access_token() -> str:
    """Acquire a token silently if cached, else prompt via device code flow."""
    app = _get_msal_app()
    accounts = app.get_accounts()
    result = None

    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        logger.info("🔐 No cached token — starting device code flow...")
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to initiate device flow: {flow}")
        print("\n" + "=" * 60)
        print(flow["message"])  # e.g. "Go to https://microsoft.com/devicelogin and enter code ABCXYZ"
        print("=" * 60 + "\n")
        result = app.acquire_token_by_device_flow(flow)

    _save_token_cache()

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    return result["access_token"]


# ── Download helpers ──────────────────────────────────────────────────────────

def _encode_sharing_url(url: str) -> str:
    """Encode a OneDrive sharing URL into a Graph API shareId (u!base64)."""
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{encoded}"


def _download_from_onedrive() -> bytes:
    """Download the Excel file from OneDrive via Microsoft Graph."""
    token = _get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    # Approach A: direct user drive path
    url_a = f"{GRAPH_BASE}/users/{FILE_OWNER_EMAIL}/drive/root:/{FILE_PATH_IN_DRIVE}:/content"
    logger.info(f"📥 Fetching from OneDrive (direct path)...")

    resp = requests.get(url_a, headers=headers, allow_redirects=True, proxies=PROXIES, timeout=30)

    if resp.status_code in (403, 404):
        # Approach B: sharing URL
        logger.warning(f"Direct path returned {resp.status_code} — trying sharing URL...")
        share_id = _encode_sharing_url(FILE_SHARING_URL)
        url_b = f"{GRAPH_BASE}/shares/{share_id}/driveItem/content"
        resp = requests.get(url_b, headers=headers, allow_redirects=True, proxies=PROXIES, timeout=30)

    resp.raise_for_status()
    logger.info(f"✅ Downloaded {len(resp.content):,} bytes from OneDrive")
    return resp.content


# ── Public API ────────────────────────────────────────────────────────────────

def refresh_file_bytes() -> None:
    """Download fresh Excel bytes from OneDrive and update the in-memory cache.

    Falls back to the local file if OneDrive is unreachable.
    Call this once before force-reloading all data modules.
    """
    global _file_bytes
    try:
        _file_bytes = _download_from_onedrive()
    except Exception as exc:
        logger.error(f"❌ OneDrive download failed: {exc}")
        if LOCAL_FALLBACK_PATH.exists():
            logger.warning("⚠️  Falling back to local Excel file.")
            _file_bytes = LOCAL_FALLBACK_PATH.read_bytes()
        else:
            raise RuntimeError(
                "OneDrive download failed and no local fallback found."
            ) from exc


def get_workbook() -> openpyxl.Workbook:
    """Return a fresh openpyxl workbook from cached bytes.

    Downloads from OneDrive on the very first call (or after refresh_file_bytes()).
    Each call creates a new Workbook instance so sheets can be fully iterated.
    """
    global _file_bytes
    if _file_bytes is None:
        refresh_file_bytes()
    return openpyxl.load_workbook(io.BytesIO(_file_bytes), data_only=True, read_only=True)
