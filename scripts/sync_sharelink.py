#!/usr/bin/env python3
"""Download an Excel workbook from a OneDrive / SharePoint sharing link."""

import os
import sys
from pathlib import Path
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

import requests

TIMEOUT = 90
XLSX_MAGIC = b"PK\x03\x04"  # .xlsx is a ZIP container
USER_AGENT = "Mozilla/5.0 (compatible; onedrive-sync/1.0)"


def fail(message: str) -> None:
    print(f"ERROR: {message}", file=sys.stderr)
    sys.exit(1)


def require(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        fail(f"Missing required environment variable: {name}")
    return value


def to_direct_download(share_url: str) -> str:
    """Add the download flag SharePoint uses to return raw file bytes."""
    parts = urlparse(share_url)
    if not parts.scheme.startswith("http"):
        fail("SHARE_URL must be a full https:// link")

    query = parse_qs(parts.query, keep_blank_values=True)
    query["download"] = ["1"]
    return urlunparse(parts._replace(query=urlencode(query, doseq=True)))


def download(url: str) -> bytes:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    response = session.get(url, timeout=TIMEOUT, allow_redirects=True)
    if response.status_code != 200:
        fail(f"Download failed with HTTP {response.status_code}")

    content = response.content
    content_type = response.headers.get("Content-Type", "")

    if not content.startswith(XLSX_MAGIC):
        if "text/html" in content_type:
            fail(
                "The link returned a sign-in page instead of the file. The sharing "
                "link is probably scoped to your organization rather than to "
                "'Anyone with the link', or it has expired."
            )
        fail(f"Unexpected content type '{content_type}' ({len(content)} bytes)")

    return content


def export_csv(workbook_path: Path) -> None:
    """Write one CSV per sheet so Git can show row-level diffs."""
    try:
        import pandas as pd
    except ImportError:
        print("pandas not installed, skipping CSV export")
        return

    csv_dir = workbook_path.parent / f"{workbook_path.stem}_csv"
    csv_dir.mkdir(parents=True, exist_ok=True)

    sheets = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl")
    for sheet_name, frame in sheets.items():
        safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in sheet_name)
        frame.to_csv(csv_dir / f"{safe.strip()}.csv", index=False)
        print(f"Exported sheet '{sheet_name}' ({len(frame)} rows)")


def main() -> None:
    share_url = require("SHARE_URL")
    dest_path = Path(require("DEST_PATH"))

    content = download(to_direct_download(share_url))

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    dest_path.write_bytes(content)
    print(f"Saved {len(content)} bytes to {dest_path}")

    if os.environ.get("EXPORT_CSV", "").lower() == "true":
        export_csv(dest_path)


if __name__ == "__main__":
    main()
