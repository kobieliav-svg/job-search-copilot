#!/usr/bin/env python3
"""Keep the user's application documents filed.

  - documents for a job that was NOT submitted yet -> output/applications/<Company_Role>/
  - documents for a job the user archived -> output/Archived/<Company_Role>/
  - documents for a submitted job (status Applied / Interviewing / Offer / Rejected, or an
    applied_date) -> output/Submitted/<Company_Role>/

Files are matched to jobs by the same name base prepare_application.py gives them:
slug(company, 30) + "_" + slug(first line of title, 40). Only Word files (.docx) are moved;
nothing is ever deleted; a name clash in the target folder gets a _2, _3 ... suffix.
Old flat files already in output/Submitted/ are left where they are.

Usage:  python scripts/file_documents.py [--dry-run]
Prints JSON: {"moved": [[from, to], ...], "unmatched": [...], "dirs": {job_id: "output/..."}}
"""
import argparse
import hashlib
import json
import os
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "pipeline" / "jobs.xlsx"
OUTPUT = ROOT / "output"
PENDING_DIR = OUTPUT / "applications"
SUBMITTED_DIR = OUTPUT / "Submitted"
ARCHIVED_DIR = OUTPUT / "Archived"  # documents of jobs the user archived (kept, never deleted)
SUBMITTED = {"Applied", "Interviewing", "Offer", "Rejected"}
SUFFIX = re.compile(r"_(Resume|CoverLetter)(_v\d+|_\d+)?$", re.I)


def slug(text, n):  # identical to prepare_application.slug
    words = re.sub(r"[^A-Za-z0-9 ]+", " ", text).split()
    return "".join(w[:1].upper() + w[1:] for w in words)[:n] or "Job"


def base_of(company, title):
    return f"{slug(company or '', 30)}_{slug((title or '').split(chr(10))[0], 40)}"


def load_jobs():
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    rows = wb["Jobs"].iter_rows(values_only=True)
    hdr = [str(h or "").strip().lower() for h in next(rows)]
    ix = {n: i for i, n in enumerate(hdr)}
    jobs = {}
    for r in rows:
        url = str(r[ix["url"]] or "").strip()
        if not url:
            continue
        status = str(r[ix["status"]] or "").strip()
        b = base_of(str(r[ix["company"]] or ""), str(r[ix["title"]] or ""))
        sub = status in SUBMITTED
        arch = status == "Archived"
        jid = hashlib.sha1(url.encode("utf-8")).hexdigest()[:16]
        # several jobs can share a base (same title reposted): submitted wins
        prev = jobs.get(b)
        if prev is None or (sub and not prev["submitted"]):
            jobs[b] = {"submitted": sub, "archived": arch and not (prev and not prev.get("archived")), "ids": (prev["ids"] if prev else []) + [jid]}
        else:
            prev["ids"].append(jid)
    # Jobs the user archived are removed from the Jobs sheet and logged in Archived_Log.
    if "Archived_Log" in wb.sheetnames:
        arows = wb["Archived_Log"].iter_rows(values_only=True)
        next(arows, None)
        for r in arows:
            if not r or not r[0]:
                continue
            b = base_of(str(r[2] or ""), str(r[1] or ""))
            if b not in jobs:
                jobs[b] = {"submitted": False, "archived": True, "ids": []}
    wb.close()
    return jobs


def target_for(dir_, name):
    t = dir_ / name
    k = 2
    while t.exists():
        t = dir_ / f"{Path(name).stem}_{k}{Path(name).suffix}"
        k += 1
    return t


def file_all(dry_run=False):
    jobs = load_jobs()
    moved, unmatched, dirs = [], [], {}
    files = list(OUTPUT.glob("*.docx")) + list(PENDING_DIR.glob("*.docx")) + list(PENDING_DIR.glob("*/*.docx")) \
        + list(SUBMITTED_DIR.glob("*/*.docx")) + list(ARCHIVED_DIR.glob("*/*.docx"))
    for f in files:
        if f.name.startswith("~$"):
            continue  # Word lock file of an open document
        b = SUFFIX.sub("", f.stem)
        job = jobs.get(b)
        if not job:
            # older file names used a shorter title; fall back to the company when exactly one
            # tracked job of that company matches the start of the file name.
            co, _, ti = b.partition("_")
            best, best_n, tie = None, 0, False
            for k, v in jobs.items():
                kco, _, kti = k.partition("_")
                if kco != co:
                    continue
                n = len(os.path.commonprefix([ti.lower(), kti.lower()]))
                if n > best_n:
                    best, best_n, tie = v, n, False
                elif n == best_n:
                    tie = True
            if best is not None and best_n >= 12 and not tie:
                job = best
        if not job and f.parent != OUTPUT:
            continue  # already in a folder and no job to decide by - leave it
        if not job:
            unmatched.append(str(f.relative_to(ROOT)))  # older name or archived job: file it as not submitted
            job = {"submitted": False, "ids": []}
        dest_dir = (ARCHIVED_DIR if job.get("archived") else SUBMITTED_DIR if job["submitted"] else PENDING_DIR) / b
        if f.parent.resolve() != dest_dir.resolve():
            t = target_for(dest_dir, f.name)
            if not dry_run:
                dest_dir.mkdir(parents=True, exist_ok=True)
                try:
                    os.replace(f, t)
                except OSError as e:  # e.g. the file is open in Word
                    unmatched.append(f"{f.relative_to(ROOT)} (could not move: {e.__class__.__name__})")
                    continue
            moved.append([str(f.relative_to(ROOT)), str(t.relative_to(ROOT))])
            if not dry_run and f.parent not in (OUTPUT, PENDING_DIR, SUBMITTED_DIR, ARCHIVED_DIR):
                try:
                    f.parent.rmdir()  # leaves it if not empty or if deleting is not allowed
                except OSError:
                    pass
        for jid in job["ids"]:
            dirs[jid] = str(dest_dir.relative_to(ROOT)).replace("\\", "/")
    # Old per-preparation folders (applications/<name>_<yyyymmdd-hhmm>/) also held working files
    # (content JSON, PDFs, temp files). Those belong in output/_work/, not next to the user's documents.
    for d in list(PENDING_DIR.glob("*_????????-????")):
        if not d.is_dir() or any(x.suffix.lower() == ".docx" for x in d.iterdir()):
            continue
        rest = [x for x in d.iterdir() if x.is_file()]
        if rest and not dry_run:
            w = OUTPUT / "_work" / d.name
            w.mkdir(parents=True, exist_ok=True)
            for x in rest:
                try:
                    os.replace(x, target_for(w, x.name))
                except OSError:
                    pass
        if rest:
            moved.append([str(d.relative_to(ROOT)) + "/(working files)", "output/_work/" + d.name])
        if not dry_run:
            try:
                d.rmdir()  # only works when empty; fine if deleting is not allowed here
            except OSError:
                pass
    # a job's folder must leave "applications" once its documents moved on
    # (submitted or archived). Empty job folders are removed; where deleting is not allowed (the
    # agents' sessions cannot delete), the empty folder is moved out of sight into
    # output/_work/_empty_folders/ instead, so "applications" only shows jobs still to apply to.
    if not dry_run:
        parking = OUTPUT / "_work" / "_empty_folders"
        for parent in (PENDING_DIR, SUBMITTED_DIR):
            for d in list(parent.glob("*")):
                if not d.is_dir() or any(d.iterdir()):
                    continue
                try:
                    d.rmdir()
                except OSError:
                    try:
                        parking.mkdir(parents=True, exist_ok=True)
                        os.replace(d, target_for(parking, d.name))
                    except OSError:
                        continue
                moved.append([str(d.relative_to(ROOT)) + "/(empty folder)", "removed"])
    return {"moved": moved, "unmatched": unmatched, "dirs": dirs}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    print(json.dumps(file_all(a.dry_run), ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
