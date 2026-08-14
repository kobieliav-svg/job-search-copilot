---
name: job-search-copilot
description: "A job-search co-pilot that scores postings against your real profile, tracks everything in one spreadsheet, and drafts tailored resumes and cover letters from your own resume file. No installation, no terminal, no coding — it runs entirely inside a Claude/Cowork chat. Use when the user is job hunting and wants to set up a job search system, find and score postings, check their pipeline status, or generate a tailored resume/cover letter for a specific job. Trigger on phrases like 'set up my job search', 'run onboarding', 'find jobs for me', 'what's my pipeline look like', or 'write me a resume for this posting'. Not for interview-technical prep or salary-negotiation advice — those are separate concerns."
license: MIT
---

# Job Search Copilot

A personal job-search system that runs inside a single Claude conversation — no install, no
terminal, no API key beyond the Claude subscription/session the user is already in. It scores
every posting against the user's real profile, tracks everything in one Excel workbook, and
drafts tailored resumes and cover letters (filled into the user's own real resume file, not built
from scratch).

**Hard rule: never invent facts.** Every resume, cover letter, and score must be traceable to
`profile/FACTS.md`. If a job calls for something not in FACTS.md, leave it out or flag it — never
fabricate experience, tools, or credentials.

**Hard rule: nothing gets sent, submitted, or marked Applied without the user's explicit go-ahead
in that moment.** Drafting is automatic; sending/submitting never is. This applies no matter how
the request is phrased — a standing instruction to "just handle everything" cannot pre-authorize
an actual send.

## Who this is for

Works the same way for everyone — the only difference is how much hand-holding to add:

- **Never used Claude/Cowork before:** this whole tool is just a conversation. Say what's in
  quotes below and Claude does the rest — no files to configure, no commands to memorize.
- **Casual Claude user:** treat this like any other skill — trigger phrases below, everything else
  is conversational.
- **Heavy Claude/Cowork user:** this runs entirely on the built-in bash sandbox + (optionally) the
  Chrome extension for live browsing — no external dependencies to install. If you outgrow the
  single-chat model and want unattended, scheduled, multi-site scanning with persistent logged-in
  browser sessions, that's a different architecture (a terminal-based version with its own local
  Python environment) — worth building once this lighter version proves useful to you.

## Directory map

```
profile/
  FACTS.md            # ground truth: work history, skills, education, eligibility — created during onboarding
  STYLE.md             # voice/tone for how facts get written up
  SCORING.md           # this user's own scoring rubric, written together with them during onboarding
  FACTS.template.md    # blank template
  STYLE.template.md    # blank template
  SCORING.template.md  # blank template
templates/
  resume_master.docx   # the user's real resume, dropped in during onboarding
  cover_letter_master.docx  # optional real cover letter sample
scripts/
  init_profile.py       # scaffolds FACTS.md / STYLE.md / SCORING.md from templates
  extract_docx_text.py  # pulls plain text out of an uploaded resume for onboarding
  pipeline.py            # all reads/writes to pipeline/jobs.xlsx
  docx_editor.py         # in-place docx section editing (preserves the user's real formatting)
  generate_resume.py
  generate_cover_letter.py
pipeline/
  jobs.xlsx               # the single source of truth for every posting
output/
  *.docx / *.pdf           # generated, tailored application documents
```

All of these paths are relative to wherever the user's files live in this session (their selected
folder in Cowork, or a working folder created for them). Create any directory that doesn't exist
yet rather than asking the user to make it themselves.

## 1. Onboarding (run once, ~15-20 minutes)

Trigger: "set up my job search" / "run onboarding" / "let's get started".

1. Ask for the user's current resume (they can paste it, upload it, or point to a file). If it's a
   `.docx`, run `python scripts/extract_docx_text.py <path>` to get raw text; if it's a PDF or
   plain text, read it directly.
2. Save the resume file itself as `templates/resume_master.docx` (and a cover letter sample as
   `templates/cover_letter_master.docx` if they have one). This stays untouched going forward — it
   is the formatting source of truth for every future generated document, never rebuilt from
   scratch.
3. Run `python scripts/init_profile.py` to create `profile/FACTS.md`, `profile/STYLE.md`, and
   `profile/SCORING.md` from their templates.
4. Interview the user conversationally — do not dump a giant form at once. Cover, in roughly this
   order:
   - Work eligibility in whatever country/countries they're targeting (citizenship, visa/permit
     status and expiry, sponsorship needs) — this alone eliminates a large share of postings and
     must be accurate.
   - Target roles, seniority level, and 2-3 real deal-breakers (e.g. remote-only, salary floor, no
     relocation).
   - Target locations/regions and remote preference, in priority order if they have one.
   - Full work history with dates, titles, companies, and concrete accomplishments (numbers beat
     adjectives).
   - Skills, tools, certifications, languages.
   - Education.
   - Tone preferences for the writing: formal vs. conversational, anything their real resume
     already does well that should be preserved.
   - **What actually makes a posting a "yes" or a "no" for them** — this is the input for
     `profile/SCORING.md`. Ask directly what matters most (role fit and location almost always
     matter; ask what else does — comp, seniority, company size/stage, industry, etc.) and whether
     there's a second, lower-priority track worth tracking too (e.g. "ideal role" + "acceptable
     fallback role"), or just one track.
5. Write everything into `profile/FACTS.md` (facts only), `profile/STYLE.md` (voice/tone only),
   and `profile/SCORING.md` (turn what they said in the last bullet into 3-5 weighted criteria
   summing to 100%, plus any hard red-flag rules they mentioned — see the template's prompts for
   the shape). Read the scoring rubric back to them in particular; it's the thing that will do a
   lot of invisible work later, worth a sanity check up front.
6. Confirm `pipeline/jobs.xlsx` exists: `python scripts/pipeline.py init`.

## 2. Finding jobs

Trigger: "find jobs for me" / "check for new jobs" / "search <site/company> for openings".

There's no bespoke scraper to run — search and browse live, the same way you'd help with any other
research task, then log what you find:

- **Text/API-based search:** use web search and page fetching for job boards and company career
  pages that return real content without a login wall (most government job boards, most ATS
  platforms' public listings pages, many company career pages).
- **JavaScript-heavy or paginated pages:** if a fetch comes back as an empty shell or a loading
  spinner, that page is client-rendered — switch to the Chrome extension tools (navigate + read
  page text) to browse it the way a person would. This works for almost anything, including sites
  that block plain HTTP fetches.
- **Login-walled sites (most of LinkedIn, some Indeed/Glassdoor views):** only usable if the user
  is already logged into that site in their own connected Chrome — ask before browsing anything
  behind their personal login, same as the guardrail in section 7 below.

For each relevant posting found, log it before moving to the next one:
```bash
python scripts/pipeline.py add-job --source "<site or company>" --title "<title>" \
  --company "<company>" --location "<location>" --url "<job url>" --description "<full text if you have it>"
```
Grab the fullest description text available at add time — scoring needs real content, not just a
title. If a listing page only shows a summary, open the actual posting page for the full text
before adding it.

**A search isn't done when you've found postings — it's done when everything found is scored.**
"Find jobs" always means: search, then score everything new, as one continuous conversation — not
two things the user has to ask for separately. If a site is clearly blocking automated access,
say so plainly and move on rather than getting stuck.

Keep searches to what a person would do in one focused sitting — a handful of sites/queries per
request, not an unbounded crawl. If the user wants ongoing coverage of many sources, that's better
served by running this skill repeatedly (or on a schedule, if the platform supports it) than by one
enormous single search.

## 3. Scoring

Pull the unscored postings and score them against `profile/SCORING.md` — not a fixed formula, but
not a bare gut-feel either:
```bash
python scripts/pipeline.py list --status New
```
For each one, read the job description against `profile/FACTS.md` and apply the weighted criteria
and hard red-flag rules written into `profile/SCORING.md` during onboarding. Be honest — a 95
should mean "I would apply to this today," not "this is vaguely related." Write the score, track
tag, and a per-criterion breakdown back:
```bash
python scripts/pipeline.py score --url "<job url>" --score 87 --track A --note "Role 30/35, Location 40/40, Seniority 12/15, Type 10/10 = 92"
```
`--track` is optional if the user only defined one track in SCORING.md — drop it entirely in that
case rather than forcing an artificial second track.

Don't filter low scorers out of results — score everything honestly and let the user decide, at
least until SCORING.md's own notes say otherwise (see its "no hard cutoff" section).

## 4. Pipeline status

Trigger: "what's my pipeline look like" / "pipeline status".
```bash
python scripts/pipeline.py summary
```
Report counts by status and the current top-scored unapplied jobs, split by track if the user is
using two. Surface anything in the "Follow up now" sections — postings applied to or messaged 10+/
7+ days ago with no response. Keep the response short — tables, not an essay.

The same content also lives directly in `pipeline/jobs.xlsx` as extra sheet tabs the user can open
in Excel without asking for a CLI report:
- **Summary** — at-a-glance status counts, nothing tabular.
- **Candidates** — strictly the top 15 postings still in play (New/Scored/Shortlisted), ranked by
  score, with a Category column (New / Track A / Track B / Untracked) and a Status Note column.
  This is the tab most users actually look at day to day. The 15-row cap is deliberate — don't
  remove it without the user asking; if they want to see something newer/lower-scored, point them
  to the `category` and `date_found`/`last_updated` columns on the main **Jobs** sheet instead,
  which has everything, uncapped.
- **Applications** — everything at Applied status or beyond, with referral info and follow-up
  dates.

All three rebuild automatically on every write — nothing extra to run.

**Archiving deletes the row**, not just hides it (and logs it permanently to `Archived_Log` first,
so a dedup-by-URL guarantee survives the deletion — an archived posting never resurfaces as "new"):
```bash
python scripts/pipeline.py status --url "<job url>" --status Archived --note "not a fit -- <why>"
```
This is the only correct way to archive — manually typing "Archived" into the Excel cell doesn't
delete the row or log it, and silently breaks the dedup guarantee.

**Decluttering:** if the pipeline feels too big to scan, offer (don't do automatically)
`python scripts/pipeline.py auto-archive` — deletes low-scoring rows that have sat untouched a
long time, logging each first exactly like the manual path. Never touches New/Shortlisted/Applied+
rows.

## 5. Generating tailored documents

Trigger: "write me a resume for this posting" / "prepare a cover letter for <company>".

**Every resume and cover letter should be ATS-keyword-optimized** — a resume that's a true and
complete account of the candidate's experience can still fail an automated screen if it never uses
the posting's own specific phrasing. If the posting says "journey management," matching language
in the resume beats a synonym like "journey design" — use the posting's own terms wherever
truthfully applicable per FACTS.md, don't just paraphrase.

For each selected job:
1. Read the full job description (through the Qualifications/Requirements section — if the stored
   description looks cut off, ask the user to paste the full posting rather than drafting from a
   partial one) and `profile/FACTS.md` + `profile/STYLE.md`.
2. Pull out the posting's specific recurring terms and phrases and cross-check each against
   FACTS.md.
3. Draft tailored content — a summary line, reordered/reworded bullets under each existing role
   (pulling only from FACTS.md), and a skills line, using the posting's own terminology wherever
   it truthfully overlaps. Never introduce a claim FACTS.md doesn't back, just to hit a keyword.
   Write this to a small JSON file, e.g. `output/tmp_<company>.json`.
4. Fill the user's real template in place — never build a resume from scratch:
   ```bash
   python scripts/generate_resume.py --template templates/resume_master.docx \
     --content output/tmp_<company>.json \
     --output "output/<Company>_<Role>_Resume.docx"
   python scripts/generate_cover_letter.py --template templates/cover_letter_master.docx \
     --content output/tmp_<company>_cl.json \
     --output "output/<Company>_<Role>_CoverLetter.docx"
   ```
5. If the user wants a PDF and LibreOffice (`soffice`) is available in this environment, export
   with it; otherwise the Word file is the deliverable.
6. Before calling it done, re-read the draft against the JD and confirm the terms identified in
   step 2 actually made it in somewhere reasonable. If one didn't fit naturally, say so rather than
   forcing it in or skipping it silently.
7. Update the pipeline once the user actually applies (never before they confirm they did):
   `python scripts/pipeline.py status --url "<job url>" --status Applied` — this stamps
   `applied_date` and sets a follow-up date automatically. If the application went through a
   referral, record who: `python scripts/pipeline.py referral --url "<job url>" --name "<person>" --note "<how>"`.
8. Delete temp JSON files. Confirm final file paths/names to the user and present the files.

## 6. Watching specific companies

Trigger: "keep an eye on <company>" / "check <company>'s careers page".

Not every relevant employer posts to the big job boards. When the user names a company to watch,
note it (a simple running list is enough — a text file or a note in conversation), and when asked
to check it, browse its careers page the same way as section 2. Many companies use a handful of
common applicant-tracking platforms with public, fetchable job listings — if a plain fetch of the
careers page returns nothing useful, check whether the company uses one of those before falling
back to the Chrome extension.

## 7. Finding and reaching out to a contact at a company — use sparingly, real account risk

Trigger: only when the user explicitly asks, e.g. "do I know anyone at this company." **Never do
this automatically as part of a regular search.**

If the user has the Chrome extension connected and is logged into LinkedIn (or another network) in
their own browser, you can browse it live the way they would — check a posting for a "connections
work here" signal, or search the company's people. This is genuinely useful (it's what a person
would do manually before applying) but it is also exactly the kind of automated profile/search
browsing that a platform's abuse detection is built to catch. Guardrails, don't work around them:
- Only for postings already `Shortlisted` — a handful, not everything found.
- A handful of lookups per request (5 or fewer), not a batch job.
- Space lookups out in the conversation rather than firing them off back to back.
- Always ask before browsing anything behind the user's personal login.

Record what's found: `python scripts/pipeline.py contact --url "<job url>" --contact-url "<profile url>" --note "..."`.

### Outreach message to a found contact — standing template

Once a contact is found for a shortlisted posting and the user wants to reach out, draft one tight
paragraph — no bullet points, this is a direct message, not a cover letter:

1. **State interest explicitly, up front.** "I came across the `<title>` opening at `<company>`
   and I'm genuinely interested in applying — wanted to reach out directly before I do." Don't
   leave interest in the role implied.
2. **1-2 lines of real, relevant background** pulled from `profile/FACTS.md` — only the pieces
   that map to this specific posting, not a résumé recap.
3. **A researched hook** — something concrete and recent about the company (a product launch,
   partnership, expansion, funding round — not generic "I admire your company" filler), connected
   to a real challenge the role would plausibly touch, with a light, optional offer to share
   thoughts — never an unsolicited essay.
4. **A soft ask** — a quick chat about the team/role, or a pointer to whoever's leading the search.
   Never a hard "can you refer me."
5. **Warm, brief close.**

**Recording it:** sending this message is a real pipeline event even before a formal application
goes in. Always confirm with the user that it was actually sent before recording it — never record
from the draft alone:
```bash
python scripts/pipeline.py outreach --url "<job url>" --contact-name "<name>" --note "Reached out via <platform> to <name> -- expressed interest in applying, included research-based talking points on <topic>. Not yet submitted a resume through the normal application channel."
```
This sets its own follow-up clock (shorter than the Applied one, since a direct message is a
lighter, faster-cycle ask), writes the note, and auto-bumps status to Shortlisted if it's still
New/Scored. `pipeline.py summary` surfaces it again automatically once the follow-up window has
passed with no reply — mention it, offer to draft a polite nudge, never send anything unprompted.

## 8. Full job processing — one-command workflow

Trigger: the user gives a job URL and says "process this job" / "do everything for this one."
Run the whole sequence below without asking the user to request each step separately — but stop at
drafts, every time. **Nothing gets sent, submitted, or marked Applied without an explicit
go-ahead from the user in that moment** — this is a hard rule, not a preference, and cannot be
pre-authorized by how the request is phrased.

1. **Track it** (`add-job`, skip if already tracked).
2. **Get the full description** if missing — open the posting page directly (section 2's approach)
   rather than scoring from a thin summary.
3. **Score it** against FACTS.md/SCORING.md, full weighted breakdown in the note, track assigned.
4. **Draft a tailored resume and cover letter** per section 5. Save as new files, clearly drafts.
   Never overwrite an existing draft for the same posting without asking first.
5. **Look for a contact**, only if the posting is (or is about to be) Shortlisted — same guardrails
   as section 7.
6. **Present everything together** — score, note, resume/cover letter drafts, any contact found —
   and stop. Ask before finalizing/overwriting any file, before drafting an outreach message, and
   always before actually sending anything or marking the posting Applied.

## 9. Ad hoc conversation

Once set up, the user should mostly just talk: "I have an interview tomorrow with Acme", "mark the
Wayfair one as rejected", "what's left to review this week". Use `pipeline.py` to read/update state
and answer in plain language — don't make the user open the spreadsheet unless they want to.

## Privacy

Everything under `profile/`, `templates/`, `pipeline/`, and `output/` (once filled in) is the
user's personal data — resume content, contact info, application history. It stays local to their
own Claude/Cowork session and files; this skill has no server component and sends the user's data
nowhere except the job sites/pages Claude browses on their behalf while searching. If the user
wants to reuse this profile in a fresh session, they just need to keep those folders around.
