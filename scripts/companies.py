#!/usr/bin/env python3
"""Manage company_targets.csv from the dashboard.

The dashboard's "Companies" page writes requests (add / pause / resume / remove) to the page's
database; the dashboard agent saves them to a JSON file and runs:

    python scripts/companies.py apply --file ops.json
        ops.json: [{"id": "...", "op": "add|pause|resume|remove", "name": "...",
                    "url": "...", "city": "...", "tier": "1|2|3", "note": "..."}]
        prints:   {"ok": true, "results": [{"id", "ok", "message", "ats_type"}]}

    python scripts/companies.py list      # the list as JSON (used by dashboard_export.py)

Rules:
- add: the site type (ats_type) is detected from the link. A company already on the list
  (same name, or same careers link) is refused with a message instead of added twice.
- pause / resume: sets the "paused" column; scan_companies.py skips paused rows.
- remove: the row is moved to company_targets_removed.csv (kept, with the date), never lost.
- The file is written to a temporary file, re-read to check it, then swapped in, so a
  failure part-way never leaves company_targets.csv half written.
"""
import argparse
import csv
import datetime
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urlparse, parse_qs

ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = Path(os.environ.get("COMPANY_TARGETS", ROOT / "company_targets.csv"))
REMOVED = CSV_PATH.with_name(CSV_PATH.stem + "_removed.csv")
FIELDS = ["company", "city", "tier", "career_url", "ats_type", "notes", "paused"]
TYPE_LABEL = {
    "workday": "Workday", "sfcsb": "SuccessFactors", "greenhouse": "Greenhouse", "lever": "Lever",
    "ashby": "Ashby", "getro": "Getro", "wp_job_manager": "WordPress job board", "generic": "general page reading",
}


def _s(v):
    return "" if v is None else str(v).strip()


def load():
    if not CSV_PATH.exists():
        return []
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        rows = []
        for r in csv.DictReader(f):
            extra = r.pop(None, None)
            if extra:  # a badly quoted row: keep the overflow in notes rather than lose it
                r["notes"] = _s(r.get("notes")) + "," + ",".join(extra)
            rows.append({k: _s(r.get(k)) for k in FIELDS})
        return rows


def save(rows, path=None):
    path = Path(path or CSV_PATH)
    tmp = path.with_name(path.name + ".new")
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)
    with open(tmp, newline="", encoding="utf-8") as f:
        check = list(csv.DictReader(f))
    if len(check) != len(rows) or any(None in r for r in check):
        raise RuntimeError("check after writing failed; the original file was not changed")
    os.replace(tmp, path)


def detect_type(url):
    p = urlparse(url)
    host = p.netloc.lower()
    q = parse_qs(p.query)
    if "myworkdayjobs.com" in host:
        return "workday"
    if "greenhouse.io" in host:
        return "greenhouse"
    if host.endswith("lever.co"):
        return "lever"
    if "ashbyhq.com" in host:
        return "ashby"
    if p.path.rstrip("/").lower().endswith("/search") and ("q" in q or "locationsearch" in q):
        return "sfcsb"  # SuccessFactors career sites: <host>/search/?q=...&locationsearch=...
    return "generic"


def _norm_url(u):
    p = urlparse(u.strip().lower())
    return (p.netloc.replace("www.", ""), p.path.rstrip("/"))


def _find(rows, name):
    n = name.strip().lower()
    for i, r in enumerate(rows):
        if r["company"].strip().lower() == n:
            return i
    return -1


def apply_ops(ops):
    rows = load()
    removed, results = [], []
    today = datetime.date.today().isoformat()
    for op in ops:
        oid, kind, name = op.get("id", ""), _s(op.get("op")).lower(), _s(op.get("name"))
        res = {"id": oid, "ok": False, "message": "", "ats_type": ""}
        if not name:
            res["message"] = "No company name."
        elif kind == "add":
            url = _s(op.get("url"))
            tier = _s(op.get("tier")) or "2"
            if not re.match(r"^https?://[^/\s]+\.[^/\s]+", url):
                res["message"] = "The careers link is not a full web address (it must start with https://)."
            elif _find(rows, name) >= 0:
                res["message"] = f"{name} is already on the list."
            elif any(_norm_url(r["career_url"]) == _norm_url(url) for r in rows if r["career_url"]):
                other = next(r["company"] for r in rows if r["career_url"] and _norm_url(r["career_url"]) == _norm_url(url))
                res["message"] = f"This careers link is already on the list (as {other})."
            else:
                t = detect_type(url)
                note = _s(op.get("note"))
                rows.append({
                    "company": name, "city": _s(op.get("city")), "tier": tier if tier in ("1", "2", "3") else "2",
                    "career_url": url, "ats_type": t,
                    "notes": f"Added {today} by the user from the dashboard." + (f" {note}" if note else ""),
                    "paused": "",
                })
                res.update(ok=True, ats_type=t, message=f"Added. Site type: {TYPE_LABEL.get(t, t)}. First scan tonight.")
        elif kind in ("pause", "resume", "remove"):
            i = _find(rows, name)
            if i < 0:
                res["message"] = f"{name} is not on the list."
            elif kind == "pause":
                rows[i]["paused"] = "yes"
                res.update(ok=True, message="Paused. It will be skipped in the nightly scan.")
            elif kind == "resume":
                rows[i]["paused"] = ""
                res.update(ok=True, message="Active again from tonight's scan.")
            else:
                r = rows.pop(i)
                r["notes"] = f"[Removed {today} from the dashboard] " + r["notes"]
                removed.append(r)
                res.update(ok=True, message="Removed. A copy is kept in company_targets_removed.csv.")
        else:
            res["message"] = f"Unknown request type: {kind}"
        results.append(res)
    if any(r["ok"] for r in results):
        if removed:
            old = []
            if REMOVED.exists():
                with open(REMOVED, newline="", encoding="utf-8") as f:
                    old = [{k: _s(r.get(k)) for k in FIELDS} for r in csv.DictReader(f)]
            save(old + removed, REMOVED)
        save(rows)
    return results


def as_list():
    return [{
        "name": r["company"], "city": r["city"], "tier": r["tier"], "url": r["career_url"],
        "type": r["ats_type"] or "generic", "paused": r["paused"].lower() in ("yes", "true", "1"),
    } for r in load() if r["ats_type"] != "directory"]


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("apply"); a.add_argument("--file", required=True)
    sub.add_parser("list")
    args = ap.parse_args()
    if args.cmd == "list":
        print(json.dumps(as_list(), ensure_ascii=False))
        return 0
    ops = json.loads(Path(args.file).read_text(encoding="utf-8"))
    try:
        results = apply_ops(ops if isinstance(ops, list) else [ops])
    except (PermissionError, RuntimeError, OSError) as e:
        print(json.dumps({"ok": False, "error": str(e)}))
        return 1
    print(json.dumps({"ok": True, "results": results}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
