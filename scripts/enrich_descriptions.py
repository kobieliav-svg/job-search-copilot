#!/usr/bin/env python3
"""Backfill the pipeline's description column by visiting each job's own URL.

Why this exists: the scanners only read search-results list cards, which never carry
the full job description -- just title/company/location. Scoring (SKILL.md section 3)
needs real description text to compare against profile/FACTS.md, so this is a required
step between scanning and scoring, not optional polish.

Run this AFTER scanning and BEFORE asking Claude to score anything:
    python scripts/enrich_descriptions.py
    python scripts/enrich_descriptions.py --source jobbank --limit 20

Each source gets its own best-effort extraction (page text minus obvious nav/boilerplate
is the fallback for anything not specifically handled). Sites change their markup, same
caveat as every other scraper in this project -- if a source stops picking up text,
check the SELECTOR comment for that source and adjust.
"""
import argparse
import time

import pipeline
from setup_browser import browser_context

# How many characters of page text to keep. Full pages include a lot of site chrome;
# this is intentionally generous rather than trying to perfectly isolate the JD.
# Generous on purpose: a cut-off posting loses its requirements and pay sections, which matter
# most for scoring and for the resume key words.
MAX_CHARS = 20000




def _extract_jobbank(page) -> str:
    # SELECTOR: Job Bank's posting detail body
    for sel in ["#jobPostingDetail", "div.job-posting-detail", "main"]:
        loc = page.locator(sel).first
        if loc.count():
            try:
                return loc.inner_text().strip()[:MAX_CHARS]
            except Exception:
                continue
    return ""



def _extract_generic(page) -> str:
    """Company-site fallback: just grab the main content area's visible text."""
    for sel in ["main", "article", "body"]:
        loc = page.locator(sel).first
        if loc.count():
            try:
                return loc.inner_text().strip()[:MAX_CHARS]
            except Exception:
                continue
    return ""


# Which browser profile and extractor to use per source. Everything else (company sites,
# job-alert links, sources the user added) uses the general page reader without any login.
SOURCE_CONFIG = {
    "jobbank": ("jobbank", _extract_jobbank),
}


def _profile_and_extractor(source: str):
    if source in SOURCE_CONFIG:
        return SOURCE_CONFIG[source]
    if source.startswith("company:"):
        return "companies", _extract_generic
    return "companies", _extract_generic  # unknown source -- best effort, no login


def enrich(jobs, limit: int, delay: float) -> int:
    updated = 0
    pending = {}  # save every 20 descriptions instead of after each one
    # Group by which browser profile they need, so we don't reopen a context per job.
    by_profile = {}
    for job in jobs[:limit]:
        profile_name, extractor = _profile_and_extractor(job["source"])
        by_profile.setdefault(profile_name, []).append((job, extractor))

    for profile_name, items in by_profile.items():
        with browser_context(profile_name, headless=True) as context:
            page = context.new_page()
            for job, extractor in items:
                print(f"Fetching description: {job['title']} @ {job['company']}...")
                try:
                    page.goto(job["url"], wait_until="domcontentloaded", timeout=20000)
                    page.wait_for_timeout(1500)
                    if extractor is _extract_generic:
                        # company ATS pages (Dayforce, SuccessFactors, iCIMS, D2L...) render the
                        # posting with JavaScript after the first load -- give them time to finish.
                        try:
                            page.wait_for_load_state("networkidle", timeout=10000)
                        except Exception:
                            pass
                        page.wait_for_timeout(1500)
                    text = extractor(page)
                except Exception as e:
                    print(f"  Failed: {e}")
                    continue
                if not text:
                    print("  No text extracted (selectors may need updating for this site).")
                    continue
                pending[job["url"]] = text
                updated += 1
                if len(pending) >= 20:
                    pipeline.update_descriptions(pending)
                    pending.clear()
                time.sleep(delay)
    if pending:
        pipeline.update_descriptions(pending)
    return updated


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--source", default=None, help="Only enrich rows whose source starts with this.")
    ap.add_argument("--limit", type=int, default=500)  # was 50; one scan adds 300+ jobs
    ap.add_argument("--delay", type=float, default=1.0)
    ap.add_argument("--days", type=int, default=0, help="Only jobs found today (0) or in the last N days.")
    args = ap.parse_args()

    jobs = pipeline.jobs_missing_description(source_filter=args.source)
    # Jobs that still matter go first (New, then Shortlisted, then Scored; newest first
    # within each), and closed rows (Archived/Rejected) are skipped entirely.
    rank = {"New": 0, "Shortlisted": 1, "Interviewing": 2, "Applied": 3, "Scored": 4}
    jobs = [j for j in jobs if str(j.get("status", "")) in rank]
    # Read descriptions only for jobs found today (plus any job the user is actively working on),
    # so old rows that never yield a description are not re-read every night.
    from datetime import date, timedelta
    since = (date.today() - timedelta(days=args.days)).isoformat()
    active = {"Shortlisted", "Interviewing", "Applied"}
    jobs = [j for j in jobs if str(j.get("status")) in active or str(j.get("date_found") or "")[:10] >= since]
    # never spend time reading postings whose title is clearly outside
    # the user's areas (title_filter.py sets those aside before this step).
    from title_filter import relevant
    jobs = [j for j in jobs if relevant(j.get("title"))]
    jobs.sort(key=lambda j: str(j.get("date_found") or ""), reverse=True)
    jobs.sort(key=lambda j: rank[str(j.get("status"))])
    if not jobs:
        print("Nothing to enrich -- every tracked job already has a description.")
        return 0

    print(f"Enriching {min(len(jobs), args.limit)} of {len(jobs)} job(s)...")
    updated = enrich(jobs, args.limit, args.delay)
    print(f"Done. Filled in descriptions for {updated} job(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
