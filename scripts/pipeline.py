#!/usr/bin/env python3
"""All reads/writes to the single-workbook job pipeline: pipeline/jobs.xlsx.

This is the source of truth the whole skill reads from. One row per posting,
deduped by URL. Designed to be called both as a CLI (from SKILL.md instructions)
and imported as a module (from the scanner scripts).

Statuses (in order): New -> Scored -> Shortlisted -> Applied -> Interviewing -> Offer / Rejected -> Archived
"""
import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.formatting.rule import ColorScaleRule

ROOT = Path(__file__).resolve().parent.parent
XLSX_PATH = ROOT / "pipeline" / "jobs.xlsx"
SHEET_NAME = "Jobs"

COLUMNS = [
    "source",       # linkedin / indeed / jobbank / glassdoor / manual
    "title",
    "company",
    "location",
    "score",        # 0-100, filled in by Claude after reading the posting, per profile/SCORING.md
    "score_note",   # per-criterion rationale for the score (see profile/SCORING.md)
    "category",     # DERIVED, not hand-edited -- recomputed from status+track on every save (see
                     # _refresh_category_column): "New", "Track A", "Track B", or "Untracked". Lets
                     # the user spot new/track at a glance while scrolling the full Jobs sheet,
                     # without switching to the Candidates tab. Added 2026-08-12.
    "url",          # dedupe key
    "date_found",
    "description",
    "status",       # pipeline stage -- one of STATUSES below
    "status_note",  # free-text detail about that stage (e.g. "phone screen Aug 20", "rejected, no reason given")
    "last_updated",
    "track",        # A = senior Ops/CX/Systems-Implementation leadership, B = Coordinator/Office-Manager admin
    "applied_date",   # set automatically the day status first becomes Applied
    "follow_up_date", # applied_date + FOLLOW_UP_DAYS; cleared once the status moves past Applied
    "outreach_sent_date",       # set by set_outreach() -- the day a networking/outreach message
                                 # (LinkedIn message to a found contact, etc.) was actually sent
    "outreach_follow_up_date",  # outreach_sent_date + OUTREACH_FOLLOW_UP_DAYS; cleared once the
                                 # status moves to Applied/Interviewing/Offer/Rejected (the outreach
                                 # did its job either way, no need to nudge on it anymore)
    "linkedin_contact_url",  # set by find_linkedin_connections.py -- scoped, opt-in, see SKILL.md
    "linkedin_contact_note", # e.g. "confirmed 1st-degree connection" vs "people-search result, not a confirmed connection"
    "referral_name",  # who the application was actually sent through, if anyone -- blank if none
    "referral_note",  # e.g. "forwarded resume directly to hiring manager, former colleague"
]

# 'remote' and 'posted_date' were dropped (2026-08-11, at the user's request -- low signal, not
# worth the column space). 'score'/'score_note' were moved right after 'location' for the same
# reason: they're the columns the user actually scans first. See _migrate_schema() -- any COLUMNS
# change (add, remove, or reorder) is handled safely by rebuilding the sheet from name-keyed data,
# not by touching cells positionally.

STATUSES = [
    "New", "Scored", "Shortlisted", "Applied",
    "Interviewing", "Offer", "Rejected", "Archived",
]

FOLLOW_UP_DAYS = 10  # how many calendar days to wait before flagging an Applied row for follow-up
OUTREACH_FOLLOW_UP_DAYS = 7  # shorter than FOLLOW_UP_DAYS -- a networking message is a lighter,
                              # faster-cycle ask than a formal application, so a nudge makes sense sooner

ARCHIVE_LOG_SHEET = "Archived_Log"
# 'Archived' postings are DELETED from the Jobs sheet outright (2026-08-12, at the user's explicit
# request -- rejected postings shouldn't sit in the file at all, not just be hidden). To keep the
# dedup-by-URL guarantee ("an archived posting never resurfaces as New in a future scan") working
# even though the row itself is gone, every archived URL is first logged here, permanently and
# append-only -- this sheet is NEVER rebuilt/cleared like Summary/Candidates/Applications are.
# add_job() checks this log in addition to the live Jobs rows before adding anything.
# NOTE: this only applies to 'Archived'. 'Rejected' (an actual application outcome) is NOT
# deleted -- that's real history the Applications tab and follow-up tracking still need.


def _ensure_archive_log(wb):
    if ARCHIVE_LOG_SHEET not in wb.sheetnames:
        log_ws = wb.create_sheet(ARCHIVE_LOG_SHEET)
        headers = ["url", "title", "company", "note", "date_archived"]
        for i, h in enumerate(headers, start=1):
            cell = log_ws.cell(row=1, column=i, value=h)
            cell.font = Font(bold=True)
        widths = [45, 45, 22, 60, 14]
        for i, w in enumerate(widths, start=1):
            log_ws.column_dimensions[get_column_letter(i)].width = w
    return wb[ARCHIVE_LOG_SHEET]


def _archived_urls(wb) -> set:
    log_ws = _ensure_archive_log(wb)
    urls = set()
    for row in range(2, log_ws.max_row + 1):
        u = log_ws.cell(row=row, column=1).value
        if u:
            urls.add(u)
    return urls


def _log_archived(wb, url, title, company, note):
    log_ws = _ensure_archive_log(wb)
    row = log_ws.max_row + 1
    log_ws.cell(row=row, column=1, value=url)
    log_ws.cell(row=row, column=2, value=title)
    log_ws.cell(row=row, column=3, value=company)
    log_ws.cell(row=row, column=4, value=note)
    log_ws.cell(row=row, column=5, value=date.today().isoformat())


def init() -> Workbook:
    XLSX_PATH.parent.mkdir(exist_ok=True)
    if XLSX_PATH.exists():
        return load_workbook(XLSX_PATH)
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    for i, col in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=i, value=col)
        cell.font = Font(bold=True)
    _apply_layout(ws)
    _save(wb)
    return wb


def _apply_layout(ws):
    widths = [10, 32, 22, 18, 8, 30, 12, 40, 14, 50, 13, 30, 14, 8, 14, 14, 14, 14, 40, 30, 20, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # Auto-filter + a colour scale on the score column, so the user can sort/filter/triage
    # directly in Excel without needing to ask for a CLI summary every time -- added 2026-08-11
    # once the row count (205+) made scrolling through the raw sheet impractical.
    last_col = get_column_letter(len(COLUMNS))
    ws.auto_filter.ref = f"A1:{last_col}1048576"
    score_col = get_column_letter(_col_index("score"))
    score_range = f"{score_col}2:{score_col}1048576"
    ws.conditional_formatting.add(
        score_range,
        ColorScaleRule(
            start_type="num", start_value=0, start_color="F8696B",     # red
            mid_type="num", mid_value=50, mid_color="FFEB84",          # yellow
            end_type="num", end_value=100, end_color="63BE7B",         # green
        ),
    )


def _migrate_schema(wb, ws) -> bool:
    """Rebuild the sheet to match the current COLUMNS list if it doesn't already match exactly.
    Existing data is extracted by column NAME (not position), so this is safe no matter how the
    schema changed since the file was created: newly added columns come back blank, removed
    columns are silently dropped, and reordered columns land in their new position -- nothing is
    ever shifted or overwritten positionally, which is what would silently corrupt real data."""
    current_header = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]
    if current_header == COLUMNS:
        return False  # already up to date

    old_name_to_idx = {name: i + 1 for i, name in enumerate(current_header) if name}
    old_rows = []
    for row in range(2, ws.max_row + 1):
        record = {name: ws.cell(row=row, column=idx).value for name, idx in old_name_to_idx.items()}
        if any(v not in (None, "") for v in record.values()):  # skip fully-blank rows
            old_rows.append(record)

    title = ws.title
    sheet_pos = wb.worksheets.index(ws)
    wb.remove(ws)
    new_ws = wb.create_sheet(title, sheet_pos)

    for i, col in enumerate(COLUMNS, start=1):
        cell = new_ws.cell(row=1, column=i, value=col)
        cell.font = Font(bold=True)
    for r, record in enumerate(old_rows, start=2):
        for i, col in enumerate(COLUMNS, start=1):
            new_ws.cell(row=r, column=i, value=record.get(col, ""))

    _apply_layout(new_ws)
    _save(wb)
    return True


def _category_for(status, track):
    """Shared by the Jobs sheet's 'category' column and the Candidates tab, so the two are always
    consistent: New status wins regardless of track, otherwise Track A/B, otherwise Untracked."""
    if status == "New":
        return "New"
    if track == "A":
        return "Track A"
    if track == "B":
        return "Track B"
    return "Untracked"


def _refresh_category_column(ws):
    """Recompute the 'category' column for every row in Jobs from its current status+track, on
    every save. This column is derived, never hand-edited -- storing it physically in Jobs (rather
    than only in the Candidates tab) lets the user see New/Track A/Track B at a glance while
    scrolling the full dataset, without switching tabs. Safe to overwrite unconditionally since
    it's fully recomputed from status+track every time, so it can never go stale."""
    cat_col = _col_index("category")
    status_col = _col_index("status")
    track_col = _col_index("track")
    url_col = _col_index("url")
    for row in range(2, ws.max_row + 1):
        if not ws.cell(row=row, column=url_col).value:
            continue  # skip blank rows
        status = ws.cell(row=row, column=status_col).value
        track = ws.cell(row=row, column=track_col).value
        ws.cell(row=row, column=cat_col, value=_category_for(status, track))


def _save(wb):
    """The one place that actually writes to disk. Every write path (add_job, score, set_status,
    etc.) calls this instead of wb.save() directly, so the category column and the
    Summary/Candidates/Applications sheets are always rebuilt from the current Jobs data and never
    go stale."""
    _refresh_category_column(wb[SHEET_NAME])
    _write_summary_sheet(wb)
    _write_candidates_sheet(wb)
    _write_applications_sheet(wb)
    wb.save(XLSX_PATH)


def _write_summary_sheet(wb):
    """Rebuild the 'Summary' tab -- just the at-a-glance header/counts, deliberately NOT mixed
    with any row-based table. Stacking multiple tables in one sheet (the original design) made it
    impossible to select/sort cleanly in Excel; the actual job lists now live in their own sheets
    (Candidates, Applications) so each sheet is a single contiguous, sortable/filterable range."""
    name = "Summary"
    if name in wb.sheetnames:
        wb.remove(wb[name])
    ws = wb.create_sheet(name, 1)

    jobs_ws = wb[SHEET_NAME]
    counts, _, total, _, _, _, follow_ups_due, outreach_follow_ups_due = summary(ws=jobs_ws)
    archived_count = len(_archived_urls(wb))

    ws.cell(row=1, column=1, value="Pipeline Summary").font = Font(bold=True, size=14)
    ws.cell(row=3, column=1, value=f"Total tracked: {total}").font = Font(bold=True)
    r = 4
    ws.cell(row=r, column=1, value="Status").font = Font(bold=True)
    ws.cell(row=r, column=2, value="Count").font = Font(bold=True)
    r += 1
    for s in STATUSES:
        if s == "Archived":
            continue  # archived rows are deleted, not counted here -- see the line below instead
        ws.cell(row=r, column=1, value=s)
        ws.cell(row=r, column=2, value=counts.get(s, 0))
        r += 1
    ws.cell(row=r, column=1, value="Archived (removed from file)")
    ws.cell(row=r, column=2, value=archived_count)
    r += 1
    r += 1
    if follow_ups_due:
        ws.cell(row=r, column=1,
                value=f"{len(follow_ups_due)} application(s) need follow-up -- see the Applications tab, "
                      f"sort/filter by Follow Up Date.").font = Font(italic=True)
        r += 1
    if outreach_follow_ups_due:
        ws.cell(row=r, column=1,
                value=f"{len(outreach_follow_ups_due)} outreach message(s) sent {OUTREACH_FOLLOW_UP_DAYS}+ days ago "
                      f"with no reply yet -- see Jobs, sort/filter by Outreach Follow Up Date.").font = Font(italic=True)
        r += 1

    ws.cell(row=r + 2, column=1,
            value="See the Candidates tab for every New/Scored/Shortlisted posting as one sortable table.").font = Font(italic=True)

    for col_letter, width in zip("AB", [55, 10]):
        ws.column_dimensions[col_letter].width = width


CANDIDATES_LIMIT = 15  # only the top N by score surface here, full stop -- see 2026-08-12 note below.


def _write_candidates_sheet(wb):
    """Rebuild the 'Candidates' tab: strictly the top CANDIDATES_LIMIT postings still in play
    (New/Scored/Shortlisted -- not yet applied, not archived), ranked by score highest-first, as
    ONE flat table with a Category column (New / Track A / Track B / Untracked). Includes
    Status Note (added 2026-08-13, at the user's request -- this is the primary tab they actually
    look at day to day, so outreach/status notes need to be visible here too, not just on Jobs).

    2026-08-12: briefly tried auto-including 'New' rows and everything touched today, uncapped, so
    a same-day scan-and-score run stayed visible here regardless of score. The user reverted that:
    they want this tab to stay a strict top-15-by-score, full stop, and want visibility into
    newly-added/newly-scored postings some other way instead -- see the 'category' column added to
    the Jobs sheet itself (_refresh_category_column) for that. Don't reintroduce an uncapped group
    here without being asked again."""
    name = "Candidates"
    if name in wb.sheetnames:
        wb.remove(wb[name])
    ws = wb.create_sheet(name, 2)

    jobs_ws = wb[SHEET_NAME]
    rows = [r for r in list_jobs(ws=jobs_ws) if r["status"] in ("New", "Scored", "Shortlisted")]

    def _sort_key(r):
        try:
            score_val = float(r["score"]) if r["score"] not in (None, "") else -1
        except (TypeError, ValueError):
            score_val = -1
        return score_val

    rows.sort(key=_sort_key, reverse=True)
    rows = rows[:CANDIDATES_LIMIT]

    fields = ("category", "title", "company", "location", "score", "status", "status_note", "url")
    for i, f in enumerate(fields, start=1):
        ws.cell(row=1, column=i, value=f.replace("_", " ").title()).font = Font(bold=True)
    for r, rec in enumerate(rows, start=2):
        ws.cell(row=r, column=1, value=_category_for(rec["status"], rec.get("track")))
        for i, f in enumerate(fields[1:], start=2):
            ws.cell(row=r, column=i, value=rec.get(f, ""))

    widths = [12, 45, 22, 22, 8, 13, 50, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(fields))
    ws.auto_filter.ref = f"A1:{last_col}{max(len(rows) + 1, 1)}"

    score_range = f"E2:E{max(len(rows) + 1, 2)}"
    ws.conditional_formatting.add(
        score_range,
        ColorScaleRule(
            start_type="num", start_value=0, start_color="F8696B",
            mid_type="num", mid_value=50, mid_color="FFEB84",
            end_type="num", end_value=100, end_color="63BE7B",
        ),
    )


def _write_applications_sheet(wb):
    """Rebuild the 'Applications' tab: every job at Applied status or beyond, with the
    referral/follow-up fields that matter once you've actually applied, so the user doesn't have
    to filter the full Jobs sheet by hand to see what's in flight."""
    name = "Applications"
    if name in wb.sheetnames:
        wb.remove(wb[name])
    ws = wb.create_sheet(name, 3)

    jobs_ws = wb[SHEET_NAME]
    rows = list_jobs(ws=jobs_ws)
    applied_rows = [r for r in rows if r["status"] in ("Applied", "Interviewing", "Offer", "Rejected")]
    # applied_date can come back as either a plain ISO string or an openpyxl datetime object
    # (Excel/older rows), depending on how it was originally written -- coerce to str for a safe,
    # still date-ordered sort rather than letting a mixed-type comparison raise TypeError.
    applied_rows.sort(key=lambda r: str(r.get("applied_date") or ""), reverse=True)

    fields = ("title", "company", "status", "applied_date", "follow_up_date",
              "referral_name", "referral_note", "status_note", "url")
    for i, f in enumerate(fields, start=1):
        ws.cell(row=1, column=i, value=f.replace("_", " ").title()).font = Font(bold=True)
    for r, rec in enumerate(applied_rows, start=2):
        for i, f in enumerate(fields, start=1):
            ws.cell(row=r, column=i, value=rec.get(f, ""))

    widths = [32, 22, 13, 14, 14, 20, 30, 30, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(fields))
    ws.auto_filter.ref = f"A1:{last_col}{max(len(applied_rows) + 1, 1)}"


def _load():
    if not XLSX_PATH.exists():
        init()
    wb = load_workbook(XLSX_PATH)
    ws = wb[SHEET_NAME]
    if _migrate_schema(wb, ws):
        ws = wb[SHEET_NAME]  # migration replaced the worksheet object -- re-fetch, don't use the stale one
    elif not ws.auto_filter.ref:
        # Schema already matches COLUMNS (no rebuild happened above), but this file predates the
        # auto-filter/colour-scale layout -- apply it once. Cheap no-op on every subsequent load
        # once auto_filter.ref is set.
        _apply_layout(ws)
        _save(wb)
    return wb, ws


def _col_index(name: str) -> int:
    return COLUMNS.index(name) + 1


def _find_row_by_url(ws, url: str):
    col = _col_index("url")
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=col).value == url:
            return row
    return None


def add_job(source, title, company, location, url, remote="", posted_date="", description=""):
    # 'remote' and 'posted_date' are still accepted (and silently dropped) so the scanner
    # scripts, which collect them from listings, don't need to be touched -- see the COLUMNS
    # comment above for why the columns themselves were removed.
    wb, ws = _load()
    existing = _find_row_by_url(ws, url)
    if existing:
        return False, existing  # already tracked, not added again
    if url in _archived_urls(wb):
        return False, "archived"  # previously reviewed and archived -- never resurfaces, even deleted
    row = ws.max_row + 1
    values = {
        "source": source,
        "title": title,
        "company": company,
        "location": location,
        "url": url,
        "date_found": date.today().isoformat(),
        "description": description,
        "score": "",
        "score_note": "",
        "status": "New",
        "status_note": "",
        "last_updated": date.today().isoformat(),
        "track": "",
        "applied_date": "",
        "follow_up_date": "",
    }
    for col_name, value in values.items():
        ws.cell(row=row, column=_col_index(col_name), value=value)
    _save(wb)
    return True, row


def score(url, value, note="", track=""):
    wb, ws = _load()
    row = _find_row_by_url(ws, url)
    if not row:
        raise ValueError(f"No job found with url: {url}")
    ws.cell(row=row, column=_col_index("score"), value=value)
    ws.cell(row=row, column=_col_index("score_note"), value=note)
    if track:
        ws.cell(row=row, column=_col_index("track"), value=track)
    if ws.cell(row=row, column=_col_index("status")).value == "New":
        ws.cell(row=row, column=_col_index("status"), value="Scored")
    ws.cell(row=row, column=_col_index("last_updated"), value=date.today().isoformat())
    _save(wb)


def set_track(url, track_value):
    """Backfill/correct just the track tag on a row that's already been scored, without
    touching its score -- useful for tagging older rows scored before the track column existed."""
    if track_value not in ("A", "B"):
        raise ValueError("track must be 'A' or 'B'")
    wb, ws = _load()
    row = _find_row_by_url(ws, url)
    if not row:
        raise ValueError(f"No job found with url: {url}")
    ws.cell(row=row, column=_col_index("track"), value=track_value)
    ws.cell(row=row, column=_col_index("last_updated"), value=date.today().isoformat())
    _save(wb)


def set_status(url, status_value, note=""):
    if status_value not in STATUSES:
        raise ValueError(f"Unknown status '{status_value}'. Must be one of: {', '.join(STATUSES)}")
    wb, ws = _load()
    row = _find_row_by_url(ws, url)
    if not row:
        raise ValueError(f"No job found with url: {url}")

    if status_value == "Archived":
        # Deleted outright, not soft-hidden -- see the ARCHIVE_LOG_SHEET note above. Log first
        # (so the dedup guarantee holds even after the row is gone), then remove the row.
        title = ws.cell(row=row, column=_col_index("title")).value
        company = ws.cell(row=row, column=_col_index("company")).value
        _log_archived(wb, url, title, company, note)
        ws.delete_rows(row, 1)
        _save(wb)
        return

    ws.cell(row=row, column=_col_index("status"), value=status_value)
    if note:
        ws.cell(row=row, column=_col_index("status_note"), value=note)

    if status_value == "Applied":
        # Set applied_date the first time (don't clobber it if this is a re-save), and
        # (re)start the follow-up clock.
        today = date.today()
        if not ws.cell(row=row, column=_col_index("applied_date")).value:
            ws.cell(row=row, column=_col_index("applied_date"), value=today.isoformat())
        ws.cell(row=row, column=_col_index("follow_up_date"),
                 value=(today + timedelta(days=FOLLOW_UP_DAYS)).isoformat())
    elif status_value in ("Interviewing", "Offer", "Rejected"):
        # A real response came in -- no longer waiting, so stop flagging it for follow-up.
        # ('Archived' returns early above -- the row's gone, nothing left to clear.)
        ws.cell(row=row, column=_col_index("follow_up_date"), value="")

    if status_value in ("Applied", "Interviewing", "Offer", "Rejected"):
        # The outreach did its job either way (led to a real application, or didn't) -- stop
        # nudging to follow up on the networking message specifically once things have moved on.
        ws.cell(row=row, column=_col_index("outreach_follow_up_date"), value="")

    ws.cell(row=row, column=_col_index("last_updated"), value=date.today().isoformat())
    _save(wb)


def set_outreach(url, contact_name, note=""):
    """Record that a networking/outreach message was actually SENT to a found contact -- distinct
    from set_contact() (just 'found someone who works there') and from set_referral() (a later,
    more definite fact that an application itself went through that person). This is the middle
    step: a message went out, and it starts its own follow-up clock (OUTREACH_FOLLOW_UP_DAYS,
    shorter than the Applied one) since a LinkedIn message is a lighter, faster-cycle ask.
    Auto-bumps status to Shortlisted if it's still New/Scored -- sending an outreach message is
    itself a sign of active pursuit. Added 2026-08-13, at the user's request: follow-up tracking
    needs to cover outreach messages, not just formal applications."""
    wb, ws = _load()
    row = _find_row_by_url(ws, url)
    if not row:
        raise ValueError(f"No job found with url: {url}")
    today = date.today()
    ws.cell(row=row, column=_col_index("outreach_sent_date"), value=today.isoformat())
    ws.cell(row=row, column=_col_index("outreach_follow_up_date"),
             value=(today + timedelta(days=OUTREACH_FOLLOW_UP_DAYS)).isoformat())
    if note:
        ws.cell(row=row, column=_col_index("status_note"), value=note)
    current_status = ws.cell(row=row, column=_col_index("status")).value
    if current_status in ("New", "Scored"):
        ws.cell(row=row, column=_col_index("status"), value="Shortlisted")
    ws.cell(row=row, column=_col_index("last_updated"), value=today.isoformat())
    _save(wb)


def set_contact(url, contact_url, note=""):
    """Record a LinkedIn contact found for this job -- written by
    find_linkedin_connections.py, which is opt-in/manual, never run automatically."""
    wb, ws = _load()
    row = _find_row_by_url(ws, url)
    if not row:
        raise ValueError(f"No job found with url: {url}")
    ws.cell(row=row, column=_col_index("linkedin_contact_url"), value=contact_url)
    ws.cell(row=row, column=_col_index("linkedin_contact_note"), value=note)
    ws.cell(row=row, column=_col_index("last_updated"), value=date.today().isoformat())
    _save(wb)


def set_referral(url, name, note=""):
    """Record that an application was actually sent through a referral, and who -- distinct from
    linkedin_contact_url/note (which is just 'someone found who works there'); this is the later,
    more definite fact that the application itself went through that person."""
    wb, ws = _load()
    row = _find_row_by_url(ws, url)
    if not row:
        raise ValueError(f"No job found with url: {url}")
    ws.cell(row=row, column=_col_index("referral_name"), value=name)
    ws.cell(row=row, column=_col_index("referral_note"), value=note)
    ws.cell(row=row, column=_col_index("last_updated"), value=date.today().isoformat())
    _save(wb)


def auto_archive_stale(min_score=30, stale_days=21):
    """Delete low-scoring, long-untouched 'Scored' rows (logging each to ARCHIVE_LOG_SHEET first,
    same as set_status(..., 'Archived')) so they stop cluttering the active views -- doesn't touch
    New rows (never scored yet) or anything already Shortlisted/Applied/beyond. Does NOT change the
    no-hard-cutoff scoring policy in profile/SCORING.md -- that's about not filtering what gets
    scored in the first place; this is about decluttering the file after the fact."""
    wb, ws = _load()
    rows = list_jobs(ws=ws)
    today = date.today()
    to_archive = []  # (row_idx, url, title, company, note)
    for r in rows:
        if r["status"] != "Scored":
            continue
        try:
            if r["score"] in (None, "") or float(r["score"]) >= min_score:
                continue
        except (TypeError, ValueError):
            continue
        last_upd = r.get("last_updated")
        if not last_upd:
            continue
        try:
            days_stale = (today - date.fromisoformat(str(last_upd))).days
        except ValueError:
            continue
        if days_stale < stale_days:
            continue
        row_idx = _find_row_by_url(ws, r["url"])
        if not row_idx:
            continue
        note = f"Auto-archived: score {r['score']} < {min_score}, inactive {days_stale}+ days"
        to_archive.append((row_idx, r["url"], r["title"], r["company"], note))

    archived = [t[2] for t in to_archive]
    # Delete highest row index first so earlier indices in the list stay valid as rows shift up.
    for row_idx, url, title, company, note in sorted(to_archive, reverse=True):
        _log_archived(wb, url, title, company, note)
        ws.delete_rows(row_idx, 1)
    if to_archive:
        _save(wb)
    return archived


def list_jobs(status_filter=None, min_score=None, track_filter=None, ws=None):
    if ws is None:
        _, ws = _load()
    rows = []
    for row in range(2, ws.max_row + 1):
        record = {col: ws.cell(row=row, column=_col_index(col)).value for col in COLUMNS}
        if status_filter and record["status"] != status_filter:
            continue
        if track_filter and record.get("track") != track_filter:
            continue
        if min_score is not None:
            try:
                if record["score"] in (None, "") or float(record["score"]) < min_score:
                    continue
            except (TypeError, ValueError):
                continue
        rows.append(record)
    return rows


def update_description(url, description):
    """Backfill the description column -- used by enrich_descriptions.py once it visits
    the actual posting page, since the scanners only see list-view cards which never
    carry the full text."""
    wb, ws = _load()
    row = _find_row_by_url(ws, url)
    if not row:
        raise ValueError(f"No job found with url: {url}")
    ws.cell(row=row, column=_col_index("description"), value=description)
    ws.cell(row=row, column=_col_index("last_updated"), value=date.today().isoformat())
    _save(wb)


def jobs_missing_description(source_filter=None):
    """Rows with a real URL but no description yet -- the candidate list for
    enrich_descriptions.py."""
    rows = list_jobs()
    out = []
    for r in rows:
        if not r.get("url"):
            continue
        if r.get("description"):
            continue
        if source_filter and not str(r.get("source", "")).startswith(source_filter):
            continue
        out.append(r)
    return out


def summary(ws=None):
    if ws is None:
        _, ws = _load()
    rows = list_jobs(ws=ws)
    counts = {s: 0 for s in STATUSES}
    for r in rows:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    def _top(candidates):
        return sorted(candidates, key=lambda r: float(r["score"]), reverse=True)[:10]

    scored_unapplied = [
        r for r in rows if r["status"] in ("New", "Scored", "Shortlisted") and r["score"] not in (None, "")
    ]
    top_unapplied = _top(scored_unapplied)
    top_track_a = _top(r for r in scored_unapplied if r.get("track") == "A")
    top_track_b = _top(r for r in scored_unapplied if r.get("track") == "B")
    top_untracked = _top(r for r in scored_unapplied if r.get("track") not in ("A", "B"))

    today = date.today()
    follow_ups_due = []
    for r in rows:
        if r["status"] != "Applied" or not r.get("follow_up_date"):
            continue
        try:
            due = date.fromisoformat(str(r["follow_up_date"]))
        except ValueError:
            continue
        if due <= today:
            follow_ups_due.append(r)
    follow_ups_due.sort(key=lambda r: str(r["follow_up_date"]))  # same mixed str/datetime risk as applied_date

    outreach_follow_ups_due = []
    for r in rows:
        if not r.get("outreach_follow_up_date"):
            continue
        try:
            due = date.fromisoformat(str(r["outreach_follow_up_date"]))
        except ValueError:
            continue
        if due <= today:
            outreach_follow_ups_due.append(r)
    outreach_follow_ups_due.sort(key=lambda r: str(r["outreach_follow_up_date"]))

    return (counts, top_unapplied, len(rows), top_track_a, top_track_b, top_untracked,
            follow_ups_due, outreach_follow_ups_due)


def _print_rows(rows, fields=("title", "company", "location", "score", "status", "url")):
    if not rows:
        print("(no rows)")
        return
    widths = {f: max(len(f), *(len(str(r.get(f, ""))) for r in rows)) for f in fields}
    header = "  ".join(f.ljust(widths[f]) for f in fields)
    print(header)
    print("-" * len(header))
    for r in rows:
        print("  ".join(str(r.get(f, "")).ljust(widths[f]) for f in fields))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("init", help="Create pipeline/jobs.xlsx if it doesn't exist yet.")

    p_add = sub.add_parser("add-job")
    p_add.add_argument("--source", required=True)
    p_add.add_argument("--title", required=True)
    p_add.add_argument("--company", required=True)
    p_add.add_argument("--location", default="")
    p_add.add_argument("--url", required=True)
    p_add.add_argument("--description", default="")

    p_score = sub.add_parser("score")
    p_score.add_argument("--url", required=True)
    p_score.add_argument("--score", required=True, type=int)
    p_score.add_argument("--note", default="")
    p_score.add_argument("--track", default="", choices=["", "A", "B"],
                          help="A = primary target track, B = secondary/backup track, as defined in profile/SCORING.md")

    p_track = sub.add_parser("track", help="Tag (or retag) a row's track without touching its score.")
    p_track.add_argument("--url", required=True)
    p_track.add_argument("--track", required=True, choices=["A", "B"])

    p_status = sub.add_parser("status")
    p_status.add_argument("--url", required=True)
    p_status.add_argument("--status", required=True, choices=STATUSES)
    p_status.add_argument("--note", default="")

    p_list = sub.add_parser("list")
    p_list.add_argument("--status", default=None, choices=STATUSES)
    p_list.add_argument("--min-score", default=None, type=float)
    p_list.add_argument("--track", default=None, choices=["A", "B"])

    p_contact = sub.add_parser("contact", help="Record a LinkedIn contact found for a job.")
    p_contact.add_argument("--url", required=True)
    p_contact.add_argument("--contact-url", required=True)
    p_contact.add_argument("--note", default="")

    p_referral = sub.add_parser("referral", help="Record that an application was sent via a referral, and by whom.")
    p_referral.add_argument("--url", required=True)
    p_referral.add_argument("--name", required=True)
    p_referral.add_argument("--note", default="")

    p_outreach = sub.add_parser("outreach", help="Record that a networking/outreach message was sent to a found contact -- starts its own follow-up clock.")
    p_outreach.add_argument("--url", required=True)
    p_outreach.add_argument("--contact-name", required=True)
    p_outreach.add_argument("--note", default="")

    p_archive = sub.add_parser("auto-archive", help="Archive low-scoring, long-inactive rows to declutter active views.")
    p_archive.add_argument("--min-score", type=float, default=30)
    p_archive.add_argument("--stale-days", type=int, default=21)

    p_desc = sub.add_parser("describe")
    p_desc.add_argument("--url", required=True)
    p_desc.add_argument("--description", required=True)

    p_missing = sub.add_parser("missing-description")
    p_missing.add_argument("--source", default=None, help="Only rows whose source starts with this (e.g. 'linkedin').")

    sub.add_parser("summary")

    args = parser.parse_args()

    if args.cmd == "init":
        init()
        print(f"Pipeline ready at {XLSX_PATH}")
    elif args.cmd == "add-job":
        added, row = add_job(
            args.source, args.title, args.company, args.location, args.url,
            description=args.description,
        )
        print(f"{'Added' if added else 'Already tracked (skipped)'} row {row}: {args.title} @ {args.company}")
    elif args.cmd == "score":
        score(args.url, args.score, args.note, track=args.track)
        print(f"Scored {args.score}{' (track ' + args.track + ')' if args.track else ''}: {args.url}")
    elif args.cmd == "track":
        set_track(args.url, args.track)
        print(f"Track -> {args.track}: {args.url}")
    elif args.cmd == "contact":
        set_contact(args.url, args.contact_url, args.note)
        print(f"Contact recorded: {args.url}")
    elif args.cmd == "referral":
        set_referral(args.url, args.name, args.note)
        print(f"Referral recorded ({args.name}): {args.url}")
    elif args.cmd == "outreach":
        set_outreach(args.url, args.contact_name, args.note)
        print(f"Outreach recorded ({args.contact_name}): {args.url} -- follow up in {OUTREACH_FOLLOW_UP_DAYS} days if no reply.")
    elif args.cmd == "auto-archive":
        archived = auto_archive_stale(args.min_score, args.stale_days)
        print(f"Archived {len(archived)} stale/low-score job(s).")
        for t in archived:
            print(f"  - {t}")
    elif args.cmd == "status":
        set_status(args.url, args.status, args.note)
        print(f"Status -> {args.status}: {args.url}")
    elif args.cmd == "list":
        rows = list_jobs(status_filter=args.status, min_score=args.min_score, track_filter=args.track)
        _print_rows(rows, fields=("title", "company", "location", "score", "track", "status", "url"))
    elif args.cmd == "describe":
        update_description(args.url, args.description)
        print(f"Description updated: {args.url}")
    elif args.cmd == "missing-description":
        rows = jobs_missing_description(source_filter=args.source)
        _print_rows(rows, fields=("source", "title", "company", "url"))
    elif args.cmd == "summary":
        counts, top, total, top_a, top_b, top_untracked, follow_ups_due, outreach_follow_ups_due = summary()
        print(f"Total tracked: {total}")
        for s in STATUSES:
            print(f"  {s:14s} {counts.get(s, 0)}")
        if follow_ups_due:
            print(f"\nFollow up now -- Applied {FOLLOW_UP_DAYS}+ days ago with no response:")
            _print_rows(follow_ups_due, fields=("title", "company", "applied_date", "follow_up_date", "url"))
        if outreach_follow_ups_due:
            print(f"\nFollow up now -- outreach message sent {OUTREACH_FOLLOW_UP_DAYS}+ days ago with no reply:")
            _print_rows(outreach_follow_ups_due, fields=("title", "company", "outreach_sent_date", "outreach_follow_up_date", "url"))
        print("\nTop unapplied -- Track A (see profile/SCORING.md for what this track means for you):")
        _print_rows(top_a)
        print("\nTop unapplied -- Track B (see profile/SCORING.md for what this track means for you):")
        _print_rows(top_b)
        if top_untracked:
            print("\nTop unapplied -- no track tagged yet (run `pipeline.py track` to tag these):")
            _print_rows(top_untracked)
    return 0


if __name__ == "__main__":
    sys.exit(main())
