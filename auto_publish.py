"""Auto-publish: generate report + upload to puppy.walmart.com.

Run manually or via Windows Task Scheduler every 12 hours.

Usage:
    uv run python auto_publish.py
"""
from __future__ import annotations

import configparser
import json
import logging
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

# ---- paths & config -------------------------------------------------------
BASE_DIR       = Path(__file__).parent
HTML_FILE      = BASE_DIR / "hr_compliance_report.html"
LOG_FILE       = BASE_DIR / "auto_publish.log"
TOKEN_CFG      = Path.home() / ".code_puppy" / "puppy.cfg"

PAGE_NAME      = "hr-compliance-report"
BUSINESS       = "general"
DESCRIPTION   = "PHL5 HR Compliance - auto-updated every 12 hours"
ACCESS_LEVEL   = "business"
UPLOAD_URL     = "https://puppy.walmart.com/api/sharing/upload"

PROXIES = {
    "http":  "http://sysproxy.wal-mart.com:8080",
    "https": "http://sysproxy.wal-mart.com:8080",
}
# ---------------------------------------------------------------------------

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
console = logging.StreamHandler(sys.stdout)
console.setLevel(logging.INFO)
logging.getLogger().addHandler(console)
logger = logging.getLogger(__name__)


def _get_token() -> str:
    cfg = configparser.ConfigParser()
    cfg.read(TOKEN_CFG)
    token = cfg.get("puppy", "puppy_token", fallback="")
    if not token:
        raise RuntimeError(
            f"No puppy_token found in {TOKEN_CFG}.\n"
            "Open Code Puppy once to refresh your login."
        )
    return token.strip()


def _generate_report() -> None:
    """Re-use export_report logic to regenerate the HTML."""
    import export_report
    logger.info("Generating fresh report from Excel data...")
    html = export_report.generate()
    HTML_FILE.write_text(html, encoding="utf-8")
    size_kb = HTML_FILE.stat().st_size // 1024
    logger.info(f"Report generated: {size_kb} KB")


def _upload(token: str) -> None:
    html = HTML_FILE.read_text(encoding="utf-8")
    payload = json.dumps({
        "name":         PAGE_NAME,
        "business":     BUSINESS,
        "html_content": html,
        "description":  DESCRIPTION,
        "access_level": ACCESS_LEVEL,
    }).encode("utf-8")

    req = urllib.request.Request(
        UPLOAD_URL,
        data=payload,
        headers={
            "Content-Type":  "application/json",
            "Authorization": f"Bearer {token}",
        },
        method="POST",
    )
    # Route through Walmart proxy
    proxy_handler = urllib.request.ProxyHandler(PROXIES)
    opener = urllib.request.build_opener(proxy_handler)

    try:
        with opener.open(req, timeout=60) as resp:
            result = json.loads(resp.read())
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        if exc.code == 401:
            raise RuntimeError(
                "Upload failed: token expired (401).\n"
                "Open Code Puppy once to refresh your login token."
            ) from exc
        raise RuntimeError(f"HTTP {exc.code}: {detail}") from exc

    url = result.get("url", "https://puppy.walmart.com/sharing/s0g0k3s/hr-compliance-report")
    logger.info(f"Uploaded successfully -> {url}")


def main() -> None:
    logger.info("=" * 50)
    logger.info(f"Auto-publish starting at {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    try:
        _generate_report()
        token = _get_token()
        _upload(token)
        logger.info("Done! Report is live at:")
        logger.info("https://puppy.walmart.com/sharing/s0g0k3s/hr-compliance-report")
    except Exception as exc:
        logger.error(f"FAILED: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
