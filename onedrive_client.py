"""Excel file loader — reads from OneDrive direct download URL or local fallback.

To enable live OneDrive sync, set the ONEDRIVE_DOWNLOAD_URL environment
variable to a direct download link from OneDrive/SharePoint.

How to get a direct download link:
  1. Open the file in OneDrive/SharePoint
  2. Click Share -> 'Anyone with the link can view'
  3. Copy the link, then replace '?e=...' with '?download=1'
  4. Set that URL as ONEDRIVE_DOWNLOAD_URL
"""
from __future__ import annotations

import io
import logging
import os
from pathlib import Path
from typing import Optional

import openpyxl

logger = logging.getLogger(__name__)

LOCAL_FILE_PATH = Path(__file__).parent / "phl5_people_dashboard.xlsx"
ONEDRIVE_DOWNLOAD_URL = os.getenv("ONEDRIVE_DOWNLOAD_URL", "")

PROXIES = {
    "http": "http://sysproxy.wal-mart.com:8080",
    "https": "http://sysproxy.wal-mart.com:8080",
}

_file_bytes: Optional[bytes] = None


def _download_from_url(url: str) -> bytes:
    """Download file bytes from a direct OneDrive/SharePoint URL."""
    import requests
    logger.info(f"Downloading Excel from OneDrive URL...")
    resp = requests.get(url, proxies=PROXIES, timeout=30, allow_redirects=True)
    resp.raise_for_status()
    logger.info(f"Downloaded {len(resp.content):,} bytes from OneDrive")
    return resp.content


def _load_local() -> bytes:
    """Read the local Excel file as bytes."""
    if not LOCAL_FILE_PATH.exists():
        raise FileNotFoundError(
            f"Local Excel file not found: {LOCAL_FILE_PATH}\n"
            "Either drop the file there or set ONEDRIVE_DOWNLOAD_URL."
        )
    logger.info(f"Loading local Excel file: {LOCAL_FILE_PATH}")
    return LOCAL_FILE_PATH.read_bytes()


def refresh_file_bytes() -> None:
    """Reload Excel bytes from OneDrive URL or local file.

    Call this before force-reloading all data modules.
    """
    global _file_bytes
    if ONEDRIVE_DOWNLOAD_URL:
        try:
            _file_bytes = _download_from_url(ONEDRIVE_DOWNLOAD_URL)
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
