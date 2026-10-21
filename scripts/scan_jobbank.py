#!/usr/bin/env python3
"""Scan Job Bank (jobbank.gc.ca), the Government of Canada's own job listing
service, and add new postings to the pipeline.

No login needed -- this is a public site and the most stable/scraper-friendly
of the built-in sources. Still, markup can change; selectors are marked "SELECTOR:".

Usage:
    python scripts/scan_jobbank.py --query "Data Analyst" --location "Calgary, AB" --max-results 40
"""
import argparse
import time
import urllib.parse

import pipeline
from setup_browser import browser_context, run_with_timeout

BASE = "https://www.jobbank.gc.ca/jobsearch/jobsearch"


def build_url(query: str, location: str) -> str:
    params = {"searchstring": query, "locationstring": location, "sort": "M"}
    return f"{BASE}?{urllib.parse.urlencode(params)}"


def scan(query: str, location: str, max_results: int) -> list:
    # collect the whole result page first and write the jobs file
    # ONCE per search. Writing it once per job (3,000+ rows) made each search exceed the 90s
    # watchdog.
    items = []
    deadline = time.monotonic() + 70
    with browser_context("jobbank", headless=True) as context:
        page = context.new_page()
        page.goto(build_url(query, location), wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        page.set_default_timeout(2000)

        # SELECTOR: result articles (article.action-buttons wraps an a.resultJobItem)
        cards = page.locator("article.action-buttons").all()
        if not cards:
            print("No job cards found -- Job Bank may have changed its markup.")
            return items

        for card in cards[:max_results]:
            if time.monotonic() > deadline:
                print(f"  Stopped after {len(items)} cards -- time budget used.")
                break
            try:
                # SELECTOR: link element carries the href; title is span.noctitle inside it
                link_el = card.locator("a.resultJobItem").first
                href = link_el.get_attribute("href") or ""
                url = href.split(";")[0]  # strip jsessionid
                if url and not url.startswith("http"):
                    url = "https://www.jobbank.gc.ca" + url

                title = card.locator("span.noctitle").first.inner_text().strip()
                company = card.locator("li.business").first.inner_text().strip()
                raw_loc = card.locator("li.location").first.inner_text().strip()
                # li.location contains a wb-inv "Location" label — remove it
                loc = raw_loc.replace("Location", "").strip()
            except Exception:
                continue

            if not url:
                continue

            items.append(dict(source="jobbank", title=title, company=company, location=loc, url=url))
    return items


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", required=True)
    ap.add_argument("--location", default="Canada")
    ap.add_argument("--max-results", type=int, default=40)
    args = ap.parse_args()

    items = run_with_timeout(scan, args.query, args.location, args.max_results)
    added = pipeline.add_jobs(items)  # outside the watchdog
    print(f"Job Bank: added {added} new job(s) to the pipeline.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
