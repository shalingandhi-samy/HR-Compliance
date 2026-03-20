"""File watcher — auto-reloads all data when the Excel file changes on disk.

Watches the local Excel file path. When a change is detected (file saved,
replaced, or synced by OneDrive), it triggers a full data reload automatically.
"""
from __future__ import annotations

import logging
import threading
import time
from pathlib import Path
from typing import Callable

from watchdog.events import FileSystemEventHandler, FileSystemEvent
from watchdog.observers import Observer

import onedrive_client

logger = logging.getLogger(__name__)

WATCH_PATH = onedrive_client.LOCAL_FILE_PATH


class _ExcelChangeHandler(FileSystemEventHandler):
    """Debounced handler that reloads data when the Excel file is modified."""

    def __init__(self, reload_fn: Callable[[], None], debounce_seconds: float = 3.0):
        self._reload_fn = reload_fn
        self._debounce = debounce_seconds
        self._timer: threading.Timer | None = None
        self._lock = threading.Lock()

    def _schedule_reload(self) -> None:
        """Debounce rapid file events (e.g. OneDrive writing in chunks)."""
        with self._lock:
            if self._timer:
                self._timer.cancel()
            self._timer = threading.Timer(self._debounce, self._do_reload)
            self._timer.daemon = True
            self._timer.start()

    def _do_reload(self) -> None:
        logger.info("Excel file changed on disk - reloading all data...")
        try:
            onedrive_client.refresh_file_bytes()
            self._reload_fn()
            logger.info("Data reloaded successfully after file change.")
        except Exception as exc:
            logger.error(f"Auto-reload failed: {exc}")

    def on_modified(self, event: FileSystemEvent) -> None:
        if Path(event.src_path).resolve() == WATCH_PATH.resolve():
            self._schedule_reload()

    def on_created(self, event: FileSystemEvent) -> None:
        if Path(event.src_path).resolve() == WATCH_PATH.resolve():
            self._schedule_reload()


def start_file_watcher(reload_fn: Callable[[], None]) -> None:
    """Start the background file watcher.

    Args:
        reload_fn: Called whenever the Excel file changes.
                   Typically this is main.scheduled_refresh.
    """
    if not WATCH_PATH.exists():
        logger.warning(
            f"Watch target not found: {WATCH_PATH} - file watcher inactive."
        )
        return

    handler = _ExcelChangeHandler(reload_fn)
    observer = Observer()
    observer.schedule(handler, str(WATCH_PATH.parent), recursive=False)
    observer.daemon = True
    observer.start()
    logger.info(f"Watching for Excel changes: {WATCH_PATH}")
