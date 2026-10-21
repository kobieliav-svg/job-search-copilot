#!/usr/bin/env python3
"""Reads profile/SETTINGS.json -- the one place for the user's personal search settings
(target titles, locations, sources, title filter, scoring weights, document rules).

Claude fills SETTINGS.json during the setup interview and updates it when the user asks
("widen the radius", "add project manager titles"). The scripts only read it.
Keys starting with "_" are notes for people and are ignored here.
"""
import json
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SETTINGS_PATH = ROOT / "profile" / "SETTINGS.json"
DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def load() -> dict:
    if not SETTINGS_PATH.exists():
        raise SystemExit("profile/SETTINGS.json is missing. Ask Claude: \"set up my job search\".")
    data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    return _strip_notes(data)


def _strip_notes(x):
    if isinstance(x, dict):
        return {k: _strip_notes(v) for k, v in x.items() if not str(k).startswith("_")}
    if isinstance(x, list):
        return [_strip_notes(v) for v in x]
    return x


def country(s=None) -> str:
    return ((s or load()).get("country") or "").strip()


def tracks(s=None) -> dict:
    """{"A": {"label":..., "titles":[...]}, ...} -- only tracks that have at least one title."""
    t = (s or load()).get("tracks") or {}
    return {k: v for k, v in t.items() if v and [x for x in v.get("titles") or [] if str(x).strip()]}


def all_titles(s=None) -> list:
    out = []
    for v in tracks(s).values():
        for x in v.get("titles") or []:
            x = str(x).strip()
            if x and x not in out:
                out.append(x)
    return out


def locations(s=None) -> list:
    """Locations in priority order: [{"place", "radius_km", "remote_only", "priority"}]."""
    locs = [l for l in (s or load()).get("locations") or [] if str(l.get("place") or "").strip()]
    return sorted(locs, key=lambda l: int(l.get("priority") or 9))


def location_words(s=None) -> list:
    """Lower-case words used to recognise a posting in the user's area (city names, country, 'remote')."""
    words = {"remote"}
    c = country(s)
    if c:
        words.add(c.lower())
    for l in locations(s):
        place = str(l.get("place") or "")
        for part in place.split(","):
            part = part.strip().lower()
            if len(part) >= 3:
                words.add(part)
    return sorted(words)


def source(name, s=None) -> dict:
    return ((s or load()).get("sources") or {}).get(name) or {}


def source_enabled(name, s=None) -> bool:
    return bool(source(name, s).get("enabled"))


def runs_today(name, s=None, today=None) -> bool:
    """True if a source with a "days" list (e.g. Job Bank) should run today. No list = every scan."""
    days = source(name, s).get("days")
    if not days:
        return True
    d = DAY_NAMES[(today or date.today()).weekday()]
    return d in days


def extra_sources(s=None) -> list:
    return [x for x in ((s or load()).get("sources") or {}).get("extra") or [] if x.get("enabled", True)]


def title_filter(s=None) -> dict:
    tf = (s or load()).get("title_filter") or {}
    return {"keep_words": [w.lower() for w in tf.get("keep_words") or [] if w],
            "exclude_words": [w.lower() for w in tf.get("exclude_words") or [] if w]}


def weights(s=None) -> dict:
    return ((s or load()).get("scoring") or {}).get("weights") or {"role_fit": 40, "location": 25, "seniority": 20, "type_and_pay": 15}


def board_min_score(s=None) -> int:
    return int(((s or load()).get("scoring") or {}).get("show_on_board_from") or 60)


def documents(s=None) -> dict:
    d = (s or load()).get("documents") or {}
    return {"resume_max_pages": int(d.get("resume_max_pages") or 2), "standing_notes": list(d.get("standing_notes") or [])}


def scan_tonight(s=None, now=None) -> bool:
    """True if tonight is a scan night in SETTINGS.json -> schedule.scan_nights. A scan that starts
    after midnight belongs to the night before (01:00 on Monday = Sunday night)."""
    from datetime import datetime, timedelta
    nights = ((s or load()).get("schedule") or {}).get("scan_nights")
    if not nights:
        return True
    night = ((now or datetime.now()) - timedelta(hours=12)).weekday()
    return DAY_NAMES[night] in nights


if __name__ == "__main__":
    import sys
    if sys.argv[1:] == ["scan-tonight"]:
        sys.exit(0 if scan_tonight() else 1)
