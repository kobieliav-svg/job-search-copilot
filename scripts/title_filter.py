#!/usr/bin/env python3
"""Title filter: set aside postings whose title is clearly outside the user's target areas,
before any time or Claude usage is spent reading them.

The word lists come from profile/SETTINGS.json -> title_filter:
  keep_words    -- if not empty, a title must contain at least one of these to be kept
  exclude_words -- a title containing any of these is set aside
Set-aside postings get status Scored, score 0 and a clear note. Nothing is deleted: the user can
move any of them back to New (on the dashboard or by asking Claude) and it is read and scored.

Usage:  python scripts/title_filter.py apply [--dry-run]
"""
import argparse
import json
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import settings

NOTE = ("Set aside by the title filter on {d}: the title is outside your target areas, so the posting "
        "was not read or scored. Move it back to New if this is wrong.")
_WORDS = None


def _words():
    global _WORDS
    if _WORDS is None:
        try:
            _WORDS = settings.title_filter()
        except SystemExit:
            _WORDS = {"keep_words": [], "exclude_words": []}
    return _WORDS


def relevant(title):
    t = (title or "").split("\n")[0].lower()
    w = _words()
    if any(x in t for x in w["exclude_words"]):
        return False
    if w["keep_words"] and not any(x in t for x in w["keep_words"]):
        return False
    return True


def apply(dry_run=False):
    import pipeline
    wb, ws = pipeline._load()
    c = {n: pipeline._col_index(n) for n in ("title", "status", "score", "score_note", "last_updated", "url")}
    today = date.today().isoformat()
    hits = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=c["status"]).value != "New" or not ws.cell(row=r, column=c["url"]).value:
            continue
        title = ws.cell(row=r, column=c["title"]).value
        if relevant(title):
            continue
        hits.append(str(title or "").split("\n")[0])
        if not dry_run:
            ws.cell(row=r, column=c["score"], value=0)
            ws.cell(row=r, column=c["score_note"], value=NOTE.format(d=today))
            ws.cell(row=r, column=c["status"], value="Scored")
            ws.cell(row=r, column=c["last_updated"], value=today)
    if hits and not dry_run:
        pipeline._save(wb)
    return hits


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["apply"])
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    hits = apply(a.dry_run)
    print(json.dumps({"ok": True, "filtered": len(hits), "dry_run": a.dry_run, "sample": hits[:15]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
