#!/usr/bin/env python3
"""The full scan in one command (this is what the nightly scheduled scan runs).

Everything it searches comes from profile/SETTINGS.json -- nothing personal is written here.
In order:
  1. Government of Canada Job Bank (if enabled, and only on the days listed in SETTINGS.json):
     every target title x every location that is not "remote only".
  2. Company career sites listed in company_targets.csv (if enabled).
  3. Extra sources the user asked Claude to add (SETTINGS.json -> sources.extra), each with its
     own small reader script in scripts/sources/.
  4. Title filter: postings whose title is clearly outside the user's areas are set aside with
     score 0 and a note, before any time is spent reading them.
  5. Description step: opens each posting found today and saves its full text, so it can be
     scored.

Job-alert emails (Gmail) are not read here: the morning agent reads them through the Gmail
connection, adds those jobs, and fills in their descriptions.

Usage:
    python scripts/run_scan.py                  # everything that is enabled
    python scripts/run_scan.py --jobbank-now    # include Job Bank even if today is not a Job Bank day
    python scripts/run_scan.py --skip-companies --skip-enrich
"""
import argparse
import subprocess
import sys
import time
from pathlib import Path

import settings

ROOT = Path(__file__).resolve().parent
PYTHON = sys.executable


def run(cmd: list[str]) -> bool:
    start = time.monotonic()
    print(f"\n[{time.strftime('%H:%M:%S')}] $ {' '.join(cmd)}", flush=True)
    result = subprocess.run(cmd)
    print(f"[{time.strftime('%H:%M:%S')}] took {time.monotonic() - start:.0f}s", flush=True)
    return result.returncode == 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--max-results", type=int, default=40, help="Most postings read per search.")
    ap.add_argument("--jobbank-now", action="store_true")
    ap.add_argument("--skip-jobbank", action="store_true")
    ap.add_argument("--skip-companies", action="store_true")
    ap.add_argument("--skip-extra", action="store_true")
    ap.add_argument("--skip-enrich", action="store_true")
    args = ap.parse_args()

    s = settings.load()
    titles = settings.all_titles(s)
    if not titles:
        raise SystemExit("No target titles in profile/SETTINGS.json yet. Ask Claude: \"set up my job search\".")
    failures = []

    # 1. Job Bank (Canada only)
    if not args.skip_jobbank and settings.source_enabled("jobbank", s):
        if args.jobbank_now or settings.runs_today("jobbank", s):
            places = [l["place"] for l in settings.locations(s) if not l.get("remote_only")]
            for query in titles:
                for place in places:
                    if not run([PYTHON, str(ROOT / "scan_jobbank.py"), "--query", query, "--location", place,
                                "--max-results", str(args.max_results)]):
                        failures.append(f"jobbank: {query} / {place}")
        else:
            days = ", ".join(settings.source("jobbank", s).get("days") or [])
            print(f"\n[{time.strftime('%H:%M:%S')}] Job Bank skipped tonight -- it runs on: {days}.", flush=True)

    # 2. Company career sites
    if not args.skip_companies and settings.source_enabled("company_sites", s):
        if not run([PYTHON, str(ROOT / "scan_companies.py"), "--days", "14"]):
            failures.append("company sites")

    # 3. Extra sources added by the user (each has scripts/sources/<script>)
    if not args.skip_extra:
        for src in settings.extra_sources(s):
            script = src.get("script")
            if not script:
                continue
            if not run([PYTHON, str(ROOT / "sources" / script)]):
                failures.append(f"source: {src.get('name') or script}")

    # 4. Title filter
    if not run([PYTHON, str(ROOT / "title_filter.py"), "apply"]):
        failures.append("title filter")

    # 5. Descriptions for today's new postings
    if not args.skip_enrich:
        if not run([PYTHON, str(ROOT / "enrich_descriptions.py")]):
            failures.append("description step")

    print("\n" + "=" * 60)
    if failures:
        print(f"Done with {len(failures)} failure(s):")
        for f in failures:
            print(f"  - {f}")
    else:
        print("Scan complete, no failures.")
    print("Next: ask Claude to score the new postings (or let the morning agent do it).")
    print("=" * 60)


STATUS_FILE = ROOT.parent / "output" / "logs" / "scan_status.json"


def _write_status(**fields):
    """The morning agent decides "is the night scan finished?" from this file. Retries because
    a sync app (e.g. OneDrive) can hold the file for a few seconds right after a write."""
    import json
    for _ in range(6):
        try:
            STATUS_FILE.parent.mkdir(parents=True, exist_ok=True)
            STATUS_FILE.write_text(json.dumps(fields), encoding="utf-8")
            return
        except OSError:
            time.sleep(5)


if __name__ == "__main__":
    _started = time.strftime("%Y-%m-%dT%H:%M:%S")
    _write_status(state="running", started_at=_started)
    try:
        main()
    finally:
        _write_status(state="finished", started_at=_started, finished_at=time.strftime("%Y-%m-%dT%H:%M:%S"))
