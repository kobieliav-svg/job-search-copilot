#!/usr/bin/env python3
"""Template: read a job board's RSS feed and add new postings to the pipeline.

Copy this file for a new source (e.g. sources/acme_board.py), set FEED_URL and SOURCE_NAME,
and add it to profile/SETTINGS.json -> sources.extra with "script": "acme_board.py".
"""
import sys
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pipeline  # noqa: E402

FEED_URL = "https://www.example.com/jobs/feed"   # <- the board's RSS feed
SOURCE_NAME = "example_board"                    # <- short name shown as the job's source
MAX_ITEMS = 60


def main() -> int:
    req = urllib.request.Request(FEED_URL, headers={"User-Agent": "Mozilla/5.0 (job-search-copilot)"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        root = ET.fromstring(resp.read())
    items = []
    for it in root.findall(".//item")[:MAX_ITEMS]:
        title = (it.findtext("title") or "").strip()
        url = (it.findtext("link") or "").strip()
        if not title or not url:
            continue
        items.append({"source": SOURCE_NAME, "title": title, "company": (it.findtext("author") or "").strip(),
                      "location": "", "url": url})
    print(f"{SOURCE_NAME}: added {pipeline.add_jobs(items)} new job(s) to the pipeline.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
