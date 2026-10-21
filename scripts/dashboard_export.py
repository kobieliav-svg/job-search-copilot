#!/usr/bin/env python3
"""Export the job pipeline for the "job search hub" dashboard page (read-only).

Reads pipeline/jobs.xlsx WITHOUT changing it and writes one JSON file,
output/dashboard/board.json, that the dashboard's agent uploads to the page:
  - counts per status, total jobs, when the Excel file last changed
  - the board: top 15 New/Scored jobs by score, plus every Shortlisted, Applied,
    Interviewing and Offer job, each as a small card.

Usage:  python scripts/dashboard_export.py
"""
import datetime
import hashlib
import json
import os
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "pipeline" / "jobs.xlsx"
OUT = ROOT / "output" / "dashboard" / "board.json"
TOP_NEW = 15
def _min_score():
    """The New column shows only jobs at or above SETTINGS.json scoring.show_on_board_from (default 60)."""
    try:
        import settings
        return settings.board_min_score()
    except SystemExit:
        return 60


NEW_MIN_SCORE = _min_score()
MIN_DESC_CHARS = 200
DESC_CHARS = 8000  # was 20000; the page's database rejects a board over 256 KB (48 cards hit 293 KB)

COLUMN_OF_STATUS = {
    "New": "new", "Scored": "new", "Shortlisted": "shortlisted", "Applied": "applied",
    "Interviewing": "interviewing", "Offer": "offer",
}


def _s(v):
    return "" if v is None else str(v).strip()


def _clip(text, n):
    text = " ".join(_s(text).split())
    return text if len(text) <= n else text[: n - 1].rstrip() + "…"


def _desc(v):
    # A job site's "… more" button label can get captured at the end of the text; it is not part of the posting.
    text = _s(v)
    for tail in ("… more", "...more", "… See more"):
        if text.endswith(tail):
            text = text[: -len(tail)].rstrip()
    return text[:DESC_CHARS]


_VIA = [  # (words in the status note or source, label) -- first match wins
    (("by email", "emailed"), "By email"),
    (("easy apply",), "LinkedIn Easy Apply"),
    (("job bank",), "Job Bank"),
    (("workday",), "Company site (Workday)"),
    (("successfactors",), "Company site (SuccessFactors)"),
    (("dayforce",), "Company site (Dayforce)"),
    (("workable",), "Company site (Workable)"),
    (("bamboohr",), "Company site (BambooHR)"),
    (("oracle",), "Company site (Oracle)"),
    (("cornerstone", "csod"), "Company site (Cornerstone)"),
    (("adp",), "Company site (ADP)"),
    (("lever",), "Company site (Lever)"),
    (("own application form",), "Company site"),
    (("recruiter", "agency"), "Through a recruiter"),
    (("applied directly via linkedin", "via linkedin"), "LinkedIn"),
    (("portal", "career site", "careers site"), "Company site"),
]


def _applied_via(note, source):
    t = (_s(note) + " " + _s(source)).lower()
    for words, label in _VIA:
        if any(w in t for w in words):
            return label
    return ""


def _date(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y-%m-%d")
    return _s(v)[:10]


def _companies(per_company):
    """The nightly company list for the dashboard's Companies page."""
    try:
        from companies import as_list
        out = as_list()
    except Exception as e:  # the board must still export if the list cannot be read
        print(json.dumps({"warning": f"company list not exported: {e}"}))
        return None
    for c in out:
        jobs, good = per_company.get(c["name"].strip().lower(), [0, 0])
        c["jobs"], c["good"] = jobs, good
    return out


SEARCH_PART_BYTES = 150_000  # the page's database takes at most 256 KB per document


def _write_search(rows):
    """Write output/dashboard/search_<i>.json: every job in the file (and in Archived_Log) as
    [title, company, status, score, url, date, archive note?], split under the size limit.
    The agents publish each part to the page's database (collection "search", doc "p<i>")."""
    gen = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    chunks, cur, size = [], [], 0
    for r in rows:
        n = len(json.dumps(r, ensure_ascii=False)) + 1
        if cur and size + n > SEARCH_PART_BYTES:
            chunks.append(cur); cur, size = [], 0
        cur.append(r); size += n
    if cur:
        chunks.append(cur)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    for i, c in enumerate(chunks):
        f = OUT.parent / f"search_{i}.json"
        tmp = f.with_suffix(".tmp")
        tmp.write_text(json.dumps({"part": i, "of": len(chunks), "generated_at": gen, "rows": c}, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, f)
    return len(chunks)


def main():
    try:  # if jobs.xlsx was damaged, put the last good copy back before reading it
        import pipeline
        pipeline._repair_if_damaged()
    except Exception as e:
        print(f"jobs.xlsx check: {e}")
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb["Jobs"]
    rows = ws.iter_rows(values_only=True)
    hdr = [_s(h).lower() for h in next(rows)]
    ix = {name: i for i, name in enumerate(hdr)}
    today = datetime.date.today().isoformat()

    counts, total, cards, per_company = {}, 0, [], {}
    search = []  # search across the whole file from the dashboard
    for r in rows:
        if r is None or all(v is None for v in r):
            continue
        total += 1
        get = lambda name: r[ix[name]] if name in ix and ix[name] < len(r) else None
        status = _s(get("status")) or "Unknown"
        try:
            _sc = int(round(float(get("score"))))
        except (TypeError, ValueError):
            _sc = None
        search.append([_clip(_s(get("title")).split("\n")[0], 80), _clip(get("company"), 50), status, _sc,
                       _s(get("url")), _date(get("date_found")) or ""])
        counts[status] = counts.get(status, 0) + 1
        src = _s(get("source"))
        if src.startswith("company:"):  # jobs found through the company list, per company (Companies page)
            k = src[len("company:"):].strip().lower()
            pc = per_company.setdefault(k, [0, 0])
            pc[0] += 1
            try:
                pc[1] += 1 if float(get("score")) >= NEW_MIN_SCORE else 0
            except (TypeError, ValueError):
                pass
        col = COLUMN_OF_STATUS.get(status)
        if not col:
            continue
        url = _s(get("url"))
        score = get("score")
        try:
            score = int(round(float(score)))
        except (TypeError, ValueError):
            score = None
        follow = _date(get("follow_up_date"))
        cards.append({
            "id": hashlib.sha1(url.encode("utf-8")).hexdigest()[:16],
            "col": col,
            "status": status,
            "title": _clip(_s(get("title")).split("\n")[0], 140),
            "company": _clip(get("company"), 80),
            "location": _clip(get("location"), 60),
            "score": score,
            "note": _clip(get("score_note"), 320),
            "url": url,
            "source": _s(get("source")),
            "date_found": _date(get("date_found")),
            "applied_date": _date(get("applied_date")),
            "status_note": _clip(get("status_note"), 160),
            "applied_via": _s(get("applied_via")) or _applied_via(get("status_note"), get("source")),
            "interview_date": _date(get("interview_date")),
            "interview_type": _s(get("interview_type")),
            "has_desc": len(_s(get("description"))) >= MIN_DESC_CHARS,
            "track": _s(get("track")),
            "desc": _desc(get("description")),
            "contact": _clip(_s(get("referral_name")) or _s(get("contact_note")), 120),
            "contact_url": _s(get("contact_url")),
            "follow_up_due": bool(status == "Applied" and follow and follow <= today),
        })

    # Top new jobs by score; the same title at the same company (re-posted ads) is shown once.
    new, seen = [], set()
    qualifying = [c for c in cards if c["col"] == "new" and c["score"] is not None and c["score"] >= NEW_MIN_SCORE]
    for c in sorted(qualifying, key=lambda c: -c["score"]):
        key = (c["title"].lower(), c["company"].lower())
        if key in seen:
            continue
        seen.add(key)
        new.append(c)
        if len(new) == TOP_NEW:
            break
    rest = [c for c in cards if c["col"] != "new"]
    rest.sort(key=lambda c: (c["col"], c["applied_date"] or "", c["score"] or 0), reverse=True)
    # the page's database takes at most 256 KB per board. Applied jobs keep only the
    # start of their description (the documents already exist); the full text stays in jobs.xlsx.
    for c in rest:
        if c["col"] == "applied" and len(c.get("desc") or "") > 2000:
            c["desc"] = c["desc"][:2000].rstrip() + " …"
    # file the Word documents (not submitted -> output/applications/<job>/,
    # submitted -> output/Submitted/<job>/) every time the board is refreshed, and tell the page
    # where each job's documents are now. Moving files never touches jobs.xlsx.
    filed = {}
    try:
        import file_documents
        filed = file_documents.file_all()
    except Exception as e:  # filing must never break the board export
        print(f"file_documents failed: {e}")
    for c in new + rest:
        d = (filed.get("dirs") or {}).get(c["id"])
        if d:
            c["docs_dir"] = d
    if "Archived_Log" in wb.sheetnames:
        on_jobs = {x[4] for x in search}
        for r in wb["Archived_Log"].iter_rows(min_row=2, values_only=True):
            if r and r[0] and r[0] not in on_jobs:
                search.append([_clip(_s(r[1]).split("\n")[0], 80), _clip(r[2], 50), "Archived", None, _s(r[0]),
                               _date(r[4]) or "", _clip(r[3], 120)])
    parts = _write_search(search)
    mod = datetime.datetime.fromtimestamp(os.path.getmtime(XLSX), datetime.timezone.utc)
    out = {
        "generated_at": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "file_modified": mod.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "jobs_total": total,
        "counts": counts,
        "top_new_limit": TOP_NEW,
        "new_min_score": NEW_MIN_SCORE,
        "new_qualifying": len({(c["title"].lower(), c["company"].lower()) for c in qualifying}),
        "cards": new + rest,
        "companies": _companies(per_company),
        "search_parts": parts,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    tmp = OUT.with_suffix(".tmp")
    tmp.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, OUT)
    print(json.dumps({"ok": True, "cards": len(out["cards"]), "bytes": OUT.stat().st_size, "path": str(OUT.relative_to(ROOT)),
                      "documents_moved": len(filed.get("moved") or [])}))


if __name__ == "__main__":
    main()
