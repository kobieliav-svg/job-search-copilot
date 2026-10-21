#!/usr/bin/env python3
"""Helper for the morning agent: the list of new jobs to score, and saving the scores.

  list   -> writes output/_work/scoring/queue.json: every New row that has a real description
            (>= 200 chars), with the fields needed to score it per profile/SCORING.md, and prints
            counts (to_score, no_description). New rows without a description are NOT scored
             -- they stay New with the "No description" tag.
  apply  -> --file scores.json, a list of {"url", "score", "track", "note"}. Writes them in ONE
            load/save of jobs.xlsx (fast, one lock window), only to rows that are still New, and
            sets them to Scored -- the same fields pipeline.score() sets.

  missing -> prints (JSON) the New rows that have NO real description, so the morning agent can
            try to read the posting itself (Chrome, read-only) for the ones that look relevant.
  setdesc -> --url U --file F: store the text in F as that row's description (same as
            enrich_descriptions.py does), so the job can then be scored normally.

Usage:
    python scripts/score_queue.py list
    python scripts/score_queue.py apply --file output/_work/scoring/scores.json
"""
import argparse
import json
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pipeline as P  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
QDIR = ROOT / "output" / "_work" / "scoring"
MIN_DESC = 200


def cmd_list():
    rows = P.list_jobs(status_filter="New")
    todo, missing = [], 0
    for r in rows:
        desc = str(r.get("description") or "")
        # Job Bank "partner site" rows only hold a stub that points to another site -- not a description.
        if len(desc) < MIN_DESC or "provided by a partner site" in desc:
            missing += 1
            continue
        todo.append({k: ("" if r.get(k) is None else str(r.get(k))) for k in
                     ("url", "source", "title", "company", "location", "date_found", "description")})
    QDIR.mkdir(parents=True, exist_ok=True)
    out = QDIR / "queue.json"
    out.write_text(json.dumps(todo, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"ok": True, "to_score": len(todo), "no_description": missing, "queue": str(out.relative_to(ROOT))}))


def _has_desc(desc):
    return len(desc) >= MIN_DESC and "provided by a partner site" not in desc


def cmd_missing():
    out = []
    for r in P.list_jobs(status_filter="New"):
        if _has_desc(str(r.get("description") or "")):
            continue
        out.append({k: ("" if r.get(k) is None else str(r.get(k))) for k in ("url", "source", "title", "company", "location", "date_found")})
    print(json.dumps(out, ensure_ascii=False))


def cmd_setdesc(url, path):
    text = Path(path).read_text(encoding="utf-8").strip()
    if len(text) < MIN_DESC:
        print(json.dumps({"ok": False, "error": "text_too_short"}))
        return
    P.update_description(url, text[:20000])
    print(json.dumps({"ok": True, "chars": len(text)}))


def cmd_apply(path):
    scores = json.loads(Path(path).read_text(encoding="utf-8"))
    wb, ws = P._load()
    done = skipped = missing = 0
    for s in scores:
        row = P._find_row_by_url(ws, s["url"])
        if not row:
            missing += 1
            continue
        if ws.cell(row=row, column=P._col_index("status")).value != "New":
            skipped += 1
            continue
        ws.cell(row=row, column=P._col_index("score"), value=int(s["score"]))
        ws.cell(row=row, column=P._col_index("score_note"), value=s.get("note", ""))
        if s.get("track") in ("A", "B"):
            ws.cell(row=row, column=P._col_index("track"), value=s["track"])
        ws.cell(row=row, column=P._col_index("status"), value="Scored")
        ws.cell(row=row, column=P._col_index("last_updated"), value=date.today().isoformat())
        done += 1
    P._save(wb)
    print(json.dumps({"ok": True, "scored": done, "not_new_anymore": skipped, "not_found": missing}))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("list")
    sub.add_parser("missing")
    a = sub.add_parser("apply"); a.add_argument("--file", required=True)
    d = sub.add_parser("setdesc"); d.add_argument("--url", required=True); d.add_argument("--file", required=True)
    args = ap.parse_args()
    if args.cmd == "list":
        cmd_list()
    elif args.cmd == "missing":
        cmd_missing()
    elif args.cmd == "setdesc":
        cmd_setdesc(args.url, args.file)
    else:
        cmd_apply(args.file)


if __name__ == "__main__":
    main()
