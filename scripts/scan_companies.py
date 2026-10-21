#!/usr/bin/env python3
"""Scan a user-maintained watchlist of specific employers (company_targets.csv at the
repo root) instead of a general job board. This covers local/regional employers that
never show up well on big job boards -- municipal governments, universities,
hospitals, and companies that mostly post directly to their own career site -- plus
gives you a place to add any company you personally want tracked every run.

How company_targets.csv works:
    company,city,tier,career_url,ats_type,notes
- career_url: the company's careers/jobs page. Leave blank if you don't have it yet --
  the scan will skip that row and print a reminder instead of guessing a URL.
- ats_type: one of "greenhouse", "lever", "ashby", "getro", "wp_job_manager", "sfcsb", "workday", "generic",
  "directory", or "unknown". greenhouse, lever, and ashby use public JSON job-board APIs
  and are the most reliable to scan. "getro" is for multi-employer aggregator boards
  (e.g. Communitech's job board) -- unlike a single-employer row, each posting there has
  its OWN employer, so the scanner reads the per-job company name off the page instead of
  using the company_targets.csv row name. "wp_job_manager" is for multi-employer job boards built on WordPress WP Job Manager
  (like getro, but it reads the board's public RSS feed -- see scan_wp_job_manager()).
  workday/generic fall back to best-effort page scraping (fragile -- company sites vary
  enormously; treat results as "worth a manual look," not gospel).
  "directory" rows (e.g. a Chamber of Commerce page) are not scanned for jobs directly --
  ask Claude to browse the page and add real employers as new rows instead (see SKILL.md).

IMPORTANT: for a single-employer row (greenhouse/lever/ashby/workday/generic), every job
found is attributed to that row's "company" value. For "getro" rows, each job's employer
is read from the page itself -- do not point a generic/unknown row at a multi-employer
aggregator page, it will mislabel every posting with the aggregator's own name instead of
the real employer.

Usage:
    python scripts/scan_companies.py --days 14
    python scripts/scan_companies.py --tier 1          # only tier-1 (highest priority) companies
"""
import argparse
import csv
import json
import re
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin, urlparse

from dateutil import parser as dateparser

import pipeline
from setup_browser import browser_context

ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = ROOT / "company_targets.csv"


def _is_recent(date_text: str, days: int) -> bool:
    if not date_text:
        return True
    try:
        # dateutil's `default` fills in whatever component the parsed string
        # is missing (e.g. a date-only string has no time), and it does so
        # via default.replace(hour=.., minute=.., ...) -- that requires a
        # datetime, not a date. Passing date.today() here throws
        # "replace() takes at most 3 keyword arguments" as soon as a real
        # date-only string shows up (confirmed live against Ashby's API).
        parsed = dateparser.parse(date_text, fuzzy=True, default=datetime.now()).date()
        return parsed >= date.today() - timedelta(days=days)
    except (ValueError, OverflowError):
        return True


def _strip_html(html: str, max_chars: int = 4000) -> str:
    """Very rough HTML-to-text: good enough for a scored-against-FACTS.md description,
    not meant to be pretty."""
    if not html:
        return ""
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_chars]


def _slug_from_url(url: str) -> str:
    """Best-effort: pull the company slug out of a greenhouse/lever/ashby URL,
    stripping any query string (e.g. ?locationId=...) first."""
    return url.split("?")[0].rstrip("/").split("/")[-1]


# Some of these public APIs (Ashby in particular, likely Cloudflare-fronted) reject
# requests with Python's default urllib User-Agent as bot traffic. A browser-like UA
# fixes it without needing a heavier HTTP library.
_API_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}


def _json_get(api_url: str):
    req = urllib.request.Request(api_url, headers=_API_HEADERS)
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read())


def scan_greenhouse(company: str, career_url: str, days: int):
    """Greenhouse job boards expose a public, unauthenticated JSON API -- no page
    scraping needed, and it doesn't change shape often."""
    slug = _slug_from_url(career_url)
    api_url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"
    try:
        data = _json_get(api_url)
    except (urllib.error.URLError, json.JSONDecodeError) as e:
        print(f"  [{company}] Greenhouse API failed ({e}); confirm the slug '{slug}' is correct.")
        return []
    jobs = []
    for job in data.get("jobs", []):
        posted = job.get("updated_at", "")
        if not _is_recent(posted, days):
            continue
        jobs.append({
            "title": job.get("title", ""),
            "url": job.get("absolute_url", ""),
            "location": (job.get("location") or {}).get("name", ""),
            "posted_date": posted,
            "description": _strip_html(job.get("content", "")),
        })
    return jobs


def scan_ashby(company: str, career_url: str, days: int):
    """Ashby job boards (jobs.ashbyhq.com/<slug>) expose a public, unauthenticated JSON
    API -- (GET .../posting-api/job-board/<slug>). No pagination,
    returns full descriptions in one call, no date filter needed at the request level so
    we filter client-side like Job Bank does."""
    slug = _slug_from_url(career_url)
    api_url = f"https://api.ashbyhq.com/posting-api/job-board/{slug}?includeCompensation=false"
    try:
        data = _json_get(api_url)
    except (urllib.error.URLError, json.JSONDecodeError) as e:
        print(f"  [{company}] Ashby API failed ({e}); confirm the slug '{slug}' is correct.")
        return []
    jobs = []
    for job in data.get("jobs", []):
        posted = job.get("publishedAt", "") or job.get("updatedAt", "")
        if not _is_recent(posted, days):
            continue
        raw_desc = job.get("descriptionPlain") or job.get("descriptionHtml") or job.get("description") or ""
        jobs.append({
            "title": job.get("title", ""),
            "url": job.get("jobUrl", "") or job.get("applyUrl", ""),
            "location": job.get("location", ""),
            "posted_date": posted,
            "description": _strip_html(raw_desc),
        })
    return jobs


def scan_lever(company: str, career_url: str, days: int):
    """Lever also exposes a public JSON API."""
    slug = _slug_from_url(career_url)
    api_url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
    try:
        data = _json_get(api_url)
    except (urllib.error.URLError, json.JSONDecodeError) as e:
        print(f"  [{company}] Lever API failed ({e}); confirm the slug '{slug}' is correct.")
        return []
    jobs = []
    for job in data:
        created = job.get("createdAt")
        posted = date.fromtimestamp(created / 1000).isoformat() if created else ""
        if not _is_recent(posted, days):
            continue
        jobs.append({
            "title": job.get("text", ""),
            "url": job.get("hostedUrl", ""),
            "location": (job.get("categories") or {}).get("location", ""),
            "posted_date": posted,
        })
    return jobs


# Nav/boilerplate link text that regularly gets caught by the generic href-contains-"job"
# heuristic but is never an actual job title. Individual postings essentially never match
# these exactly, so this cuts a lot of the "URL went to a general listing page" noise.
_NAV_PHRASES = {
    "careers", "career", "jobs", "job search", "search jobs", "view all jobs",
    "view all openings", "see all jobs", "see all openings", "browse jobs",
    "browse all jobs", "apply now", "current opportunities", "open positions",
    "all positions", "job opportunities", "employment opportunities", "learn more",
    "sign in", "log in", "register", "job alerts", "back to search results",
}


_NOT_POSTING_WORDS = ("translate", "about", "benefit", "why-", "student", "volunteer", "equity", "culture",
                      "event", "faq", "profile", "register", "login", "sign-in", "signin", "alert", "intro",
                      "what-makes", "life-at", "working-at", "work-at", "co-operative", "training", "viewalljobs",
                      "search", "/go/", "topjobs", "employment-careers", "careers-volunteering")
_SOCIAL_HOSTS = ("twitter.com", "x.com", "facebook.com", "instagram.com", "youtube.com", "linkedin.com/company")


def _looks_like_posting(url: str) -> bool:
    """Keep only links that look like one specific posting -- a /job/ or /posting/ path
    with something after it, or a job/requisition id in the query -- and drop section pages and
    social links. Cuts the "careers page navigation" rows that were ~70% of company-site results."""
    low = url.lower()
    if any(h in low for h in _SOCIAL_HOSTS):
        return False
    parsed = urlparse(low)
    path, query = parsed.path, parsed.query
    if any(w in path for w in _NOT_POSTING_WORDS):
        return False
    if re.search(r"(jobid|job_id|requisition|reqid|req_id|posting_id|vacancyid)=", query):
        return True
    m = re.search(r"/(job|jobs|posting|postings|requisition|vacancy|vacancies|opening|openings|detail)/([^/?#]+)", path)
    return bool(m and len(m.group(2)) >= 3)


def scan_generic(company: str, career_url: str, days: int):
    """Best-effort fallback for company sites not on a known ATS (includes Workday
    tenants, which vary too much per-tenant to handle generically here). Grabs anything
    that looks like a job link on the page. Expect to hand-verify results from this path --
    it will both miss real postings and occasionally pick up noise, even after filtering
    out obvious nav links below."""
    jobs = []
    with browser_context("companies", headless=True) as context:
        page = context.new_page()
        try:
            page.goto(career_url, wait_until="domcontentloaded", timeout=20000)
        except Exception as e:
            print(f"  [{company}] Could not load {career_url}: {e}")
            return []
        page.wait_for_timeout(1500)

        # SELECTOR: heuristic -- any link whose href or text suggests a job posting
        # "career" links were dropped -- they are almost always section/navigation pages
        # (about us, benefits, students, translate), never postings. See _looks_like_posting below.
        links = page.locator("a[href*='job'], a[href*='posting'], a[href*='requisition'], a[href*='vacanc']").all()
        seen = set()
        for link in links:
            try:
                title = link.inner_text().strip()
                href = link.get_attribute("href") or ""
            except Exception:
                continue
            if not title or len(title) < 6 or href in seen:
                continue
            if title.strip().lower() in _NAV_PHRASES:
                continue
            url = urljoin(career_url, href)
            if url.rstrip("/") == career_url.rstrip("/"):
                continue  # links straight back to the listing page itself, not a posting
            if not _looks_like_posting(url):
                continue
            seen.add(href)
            jobs.append({"title": title, "url": url, "location": "", "posted_date": ""})
    return jobs[:30]  # generic scraping is noisy -- cap it and let scoring filter the junk


def scan_sfcsb(company: str, career_url: str, days: int):
    """SAP SuccessFactors Career Site Builder boards (used by many cities and large employers). These sites
    render the job list at <origin>/search/?q=&searchResultView=LIST and every posting is a
    link containing "/job/". The generic scraper reads only nav links on these sites, so this
    handler goes straight to the search-results list, reads each result's title + link +
    location, and follows pagination via &startrow=N (25 per page). The general page reader
    returns only navigation links on these sites."""
    from urllib.parse import urlparse
    parsed = urlparse(career_url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    # keep a keyword / location filter from company_targets.csv (e.g. ?q=customer&locationsearch=<your province>)
    # so big employers (Rogers, TELUS) don't flood the pipeline with every retail posting.
    from urllib.parse import parse_qs, quote_plus
    qs = parse_qs(parsed.query)
    q = quote_plus((qs.get("q") or [""])[0])
    loc = (qs.get("locationsearch") or [""])[0]
    extra = f"&locationsearch={quote_plus(loc)}" if loc else ""
    jobs, seen = [], set()
    with browser_context("companies", headless=True) as context:
        page = context.new_page()
        for startrow in range(0, 200, 25):
            url = f"{origin}/search/?q={q}{extra}&searchResultView=LIST&startrow={startrow}"
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=25000)
            except Exception as e:
                print(f"  [{company}] Could not load {url}: {e}")
                break
            page.wait_for_timeout(1500)
            # SELECTOR: each posting is an <a href*="/job/"> inside the results list
            links = page.locator("a[href*='/job/']").all()
            new_on_page = 0
            for link in links:
                try:
                    title = " ".join(link.inner_text().split())
                    href = link.get_attribute("href") or ""
                except Exception:
                    continue
                if not title or not href:
                    continue
                full = urljoin(origin, href)
                if full in seen:
                    continue
                seen.add(full)
                new_on_page += 1
                jobs.append({"title": title, "url": full, "location": "", "posted_date": ""})
            if new_on_page == 0:
                break
    return jobs


_WORKDAY_TERMS = ("customer", "client", "contact centre", "contact center", "service", "operations",
                  "implementation", "vendor", "project coordinator", "office manager")
def _area_hints():
    """Words that mark a posting as being in the user's area: city names and country from
    SETTINGS.json, plus 'remote'. Empty settings = keep every location."""
    try:
        import settings
        return tuple(settings.location_words())
    except SystemExit:
        return ()


_AREA_HINTS = _area_hints()


def _workday_recent(posted: str, days: int) -> bool:
    p = (posted or "").lower()
    if "today" in p or "yesterday" in p:
        return True
    m = re.search(r"(\d+)\+?\s+days?", p)
    if m:
        return int(m.group(1)) <= days and "+" not in p
    return True


def scan_workday(company: str, career_url: str, days: int):
    """Workday tenants (<tenant>.wdN.myworkdayjobs.com/<site>) -- the job list is
    JavaScript-rendered, so instead of scraping we call Workday's own public JSON endpoint
    (POST /wday/cxs/<tenant>/<site>/jobs) once per keyword in _WORKDAY_TERMS, keep postings in
    in the user's area (SETTINGS.json locations) posted within `days`, and fetch nothing else. Not yet confirmed by a live
    run when added -- if a tenant returns errors, the scan prints it and moves on."""
    parsed = urlparse(career_url)
    host = parsed.netloc
    parts = [x for x in parsed.path.split("/") if x]
    parts = [x for x in parts if not re.fullmatch(r"[a-z]{2}-[A-Z]{2}", x)]  # drop a locale like en-US
    if not parts:
        print(f"  [{company}] Workday URL has no site name: {career_url}")
        return []
    site = parts[0]
    tenant = host.split(".")[0]
    api = f"https://{host}/wday/cxs/{tenant}/{site}/jobs"
    headers = dict(_API_HEADERS); headers["Content-Type"] = "application/json"
    jobs, seen = [], set()
    for term in _WORKDAY_TERMS:
        body = json.dumps({"appliedFacets": {}, "limit": 20, "offset": 0, "searchText": term}).encode()
        try:
            req = urllib.request.Request(api, data=body, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read())
        except Exception as e:
            print(f"  [{company}] Workday search '{term}' failed: {e}")
            continue
        for jp in data.get("jobPostings", []) or []:
            path = jp.get("externalPath") or ""
            title = (jp.get("title") or "").strip()
            locs = (jp.get("locationsText") or "")
            if not path or not title or path in seen:
                continue
            if not _workday_recent(jp.get("postedOn", ""), days):
                continue
            low = " " + locs.lower()
            if locs and _AREA_HINTS and not any(h in low for h in _AREA_HINTS) and "locations" not in low:
                continue
            seen.add(path)
            jobs.append({"title": title, "url": f"https://{host}/{site}{path}", "location": locs, "posted_date": ""})
        time.sleep(0.5)
    return jobs


def scan_wp_job_manager(company: str, career_url: str, days: int):
    """Job boards built on WordPress WP Job Manager publish a clean public RSS feed with the
    company, location and full description of every posting, which is far more reliable than
    reading HTML. Like a Getro board, each posting has its OWN employer: read it from the feed's
    job_listing:company field, never use the board's name."""
    import xml.etree.ElementTree as ET
    from urllib.parse import urlparse

    domain = urlparse(career_url).netloc
    feed_url = f"https://{domain}/?feed=job_feed"
    req = urllib.request.Request(feed_url, headers=_API_HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            raw = resp.read()
    except urllib.error.URLError as e:
        print(f"  [{company}] RSS feed failed ({e}).")
        return []

    try:
        root = ET.fromstring(raw)
    except ET.ParseError as e:
        print(f"  [{company}] RSS feed did not parse ({e}).")
        return []

    NS = {"job_listing": f"https://{domain}"}
    jobs = []
    for item in root.findall(".//item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        pub_date = (item.findtext("pubDate") or "").strip()
        if not _is_recent(pub_date, days):
            continue
        employer = (item.findtext("job_listing:company", namespaces=NS) or "").strip() or company
        location = (item.findtext("job_listing:location", namespaces=NS) or "").strip()
        desc_el = item.find("{http://purl.org/rss/1.0/modules/content/}encoded")
        desc_raw = desc_el.text if desc_el is not None else (item.findtext("description") or "")
        if not title or not link:
            continue
        jobs.append({
            "title": title, "url": link, "location": location, "posted_date": pub_date,
            "description": _strip_html(desc_raw or ""), "company": employer,
        })
    return jobs


def scan_getro(company: str, career_url: str, days: int):
    """Multi-employer aggregator boards (Getro-powered, e.g. Communitech's Work In Tech
    job board). Unlike scan_generic, this reads each posting's ACTUAL employer name off
    the page rather than mislabeling everything with the aggregator's own name.

    Typical job-board markup: job links look like
    https://<board-host>/companies/<employer-slug>/jobs/<id>-<title-slug>#content
    -- the employer slug is right there in the URL path, which is far more reliable than
    trying to DOM-walk up to a "nearby company element" (that approach was tried first and
    kept re-matching the job's own link instead of a separate company link -- Getro doesn't
    render employer name as visible card text next to the job title, it's an icon/logo).
    If a future markup change breaks the /companies/<slug>/jobs/ URL shape, employer will
    fall back to the watchlist row's own company name and print a warning -- watch for that
    in scan output rather than silently mislabeling everything."""
    import re as _re

    jobs = []
    with browser_context("companies", headless=True) as context:
        page = context.new_page()
        try:
            page.goto(career_url, wait_until="domcontentloaded", timeout=20000)
        except Exception as e:
            print(f"  [{company}] Could not load {career_url}: {e}")
            return []
        page.wait_for_timeout(2000)

        # Getro boards lazy-load more cards as you scroll; do a few scroll+wait passes so
        # more than just the first screenful of jobs gets captured.
        for _ in range(4):
            page.mouse.wheel(0, 3000)
            page.wait_for_timeout(800)

        # SELECTOR: job posting links on a Getro board live under /companies/<slug>/jobs/<id>
        job_links = page.locator("a[href*='/companies/'][href*='/jobs/']").all()
        seen = set()
        unresolved = 0
        for link in job_links:
            try:
                title = link.inner_text().strip()
                href = link.get_attribute("href") or ""
            except Exception:
                continue
            if not title or len(title) < 4 or href in seen:
                continue
            seen.add(href)
            url = href if href.startswith("http") else "https://" + career_url.split("/")[2] + href

            # Employer slug lives between /companies/ and /jobs/ in the URL path itself.
            m = _re.search(r"/companies/([^/]+)/jobs/", url)
            employer = ""
            if m:
                employer = m.group(1).replace("-", " ").title()
            if not employer:
                unresolved += 1
                employer = f"{company} (employer not extracted -- check manually)"

            jobs.append({
                "title": title, "url": url, "location": "", "posted_date": "",
                "company": employer,
            })
        if unresolved:
            print(f"  [{company}] WARNING: {unresolved} posting(s) didn't match the expected "
                  f"/companies/<slug>/jobs/ URL shape -- markup may have changed, check manually.")
    return jobs[:50]


def scan(days: int, tier_filter: int | None = None) -> int:
    if not CSV_PATH.exists():
        print(f"No {CSV_PATH} found -- nothing to scan.")
        return 0

    added = 0
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            company = row["company"].strip()
            career_url = (row.get("career_url") or "").strip()
            ats_type = (row.get("ats_type") or "generic").strip().lower()
            tier = row.get("tier", "").strip()

            if tier_filter is not None and str(tier_filter) != tier:
                continue
            if ats_type == "directory":
                continue
            if (row.get("paused") or "").strip().lower() in ("yes", "true", "1"):
                print(f"  [{company}] Paused from the dashboard -- skipped.")
                continue
            if not career_url:
                print(f"  [{company}] No career_url set in company_targets.csv -- skipped.")
                continue

            print(f"Scanning {company} ({ats_type})...")
            if ats_type == "greenhouse":
                jobs = scan_greenhouse(company, career_url, days)
            elif ats_type == "ashby":
                jobs = scan_ashby(company, career_url, days)
            elif ats_type == "lever":
                jobs = scan_lever(company, career_url, days)
            elif ats_type == "getro":
                jobs = scan_getro(company, career_url, days)
            elif ats_type == "wp_job_manager":
                jobs = scan_wp_job_manager(company, career_url, days)
            elif ats_type == "sfcsb":
                jobs = scan_sfcsb(company, career_url, days)
            elif ats_type == "workday":
                jobs = scan_workday(company, career_url, days)
            else:
                jobs = scan_generic(company, career_url, days)

            # one write per company instead of one per job (see pipeline.add_jobs).
            batch = []
            for job in jobs:
                if not job.get("url") or not job.get("title"):
                    continue
                # Getro (and any future aggregator handler) can supply the real per-job
                # employer; everything else uses the watchlist row's own company name.
                batch.append(dict(
                    source=f"company:{company}",
                    title=job["title"], company=job.get("company") or company,
                    location=job.get("location", ""), url=job["url"],
                    description=job.get("description", ""),
                ))
            try:
                added += pipeline.add_jobs(batch)
            except Exception as e:  # one company's save problem must not end the whole watchlist
                print(f"  [{company}] Could not save {len(batch)} job(s): {e}")
            time.sleep(0.5)
    return added


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=14, help="Only jobs posted within this many days.")
    ap.add_argument("--tier", type=int, default=None, help="Only scan companies with this tier value.")
    args = ap.parse_args()

    added = scan(args.days, tier_filter=args.tier)
    print(f"Company watchlist: added {added} new job(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
