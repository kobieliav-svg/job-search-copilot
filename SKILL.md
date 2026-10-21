---
name: job-search-copilot
description: "Job-search assistant: scores job postings against the user's real background, drafts tailored resumes and cover letters from their own resume file, prepares them for interviews, and tracks every job in one spreadsheet. The Full setup adds a nightly search, a morning agent and a personal dashboard. Use when the user is job hunting: 'set up my job search', 'score this job', 'prepare my application', 'prepare me for my interview', 'add <site> to my search', 'what's my pipeline look like'. It drafts only: it never sends, applies or messages anyone."
license: MIT
---

# Job Search Copilot

A personal job-search system that runs in Claude. Two levels:

- **Basic** - everything happens in chat: scoring a posting the user pastes, tailored resume and
  cover letter, interview preparation and practice, and the tracking spreadsheet.
- **Full** - adds the nightly scan on the user's Windows computer, the morning agent (job-alert
  emails, scoring, employer replies) and the personal dashboard. Needs a paid Claude plan, the
  Claude desktop app, Python, Gmail for email features, and a computer that is on at night.
  See GUIDE.md chapter 3 and section 8 below.

The user-facing guide is `GUIDE.md`. When a user gives you the guide and asks you to set things up,
follow it step by step: do every step marked [Claude] yourself, and for every step marked [You],
stop and tell them exactly what to click.

## The iron rule (applies to everything below)

This skill prepares; the user decides and acts. Never send, submit, apply, post, connect, or
message anyone on the user's behalf, in any tool, even if asked to "just send it". Produce the
draft, say where it is, and stop. Never invent facts about the user: every document and score is
built only from `profile/FACTS.md` and what the user says.

## Directory map

```
profile/
  FACTS.md, STYLE.md, SCORING.md   # who the user is, how they sound, how jobs are scored
  SETTINGS.json                     # targets, locations, sources, title filter, weights, document rules
  *.template.*                      # blank templates (init_profile.py copies them)
templates/
  resume_master.docx                # the user's own resume layout (track A)
  resume_master_B.docx              # optional second layout for track B
  cover_letter_master.docx          # cover letter layout with {{PLACEHOLDERS}}
company_targets.csv                 # employers whose career pages are scanned (Full)
pipeline/jobs.xlsx                  # every job and its status; backups/ keeps checked copies
scripts/                            # all reading and writing goes through these
  sources/                          # readers for extra sources the user added (Full)
agents/                             # the two scheduled-task prompts (Full)
dashboard/                          # the dashboard page template (Full)
output/                             # applications/, Submitted/, Archived/ hold the Word files
```

Create any missing folder yourself rather than asking the user to.

## 1. Setup interview (run once; about 15-25 minutes)

Trigger: "set up my job search", "run setup", "let's get started", or the first request in a
folder where `profile/FACTS.md` does not exist yet.

Ask in plain conversation, a few questions at a time. Never paste the whole list as a form.
Before starting, ask one thing: **Basic or Full setup?** (Basic = everything works in chat.
Full = adds the nightly scan, the morning agent and the dashboard; needs a paid Claude plan,
the Claude desktop app on Windows, Python, and a computer that is on at night.) If they are
not sure, start Basic: Full can be added later without redoing anything.

### Step 1: create the files
Run `python scripts/init_profile.py`. It creates, from the blank templates:
`profile/FACTS.md`, `profile/STYLE.md`, `profile/SCORING.md`, `profile/SETTINGS.json`,
`company_targets.csv` and `pipeline/jobs.xlsx`. Existing files are never overwritten.

### Step 2: the resume
- **They have one:** ask them to upload it (Word or PDF). For Word, run
  `python scripts/extract_docx_text.py <file>`; read PDFs directly. Save the Word file as
  `templates/resume_master.docx`; it is the layout every tailored resume is built on. If they
  only have a PDF, build the master in Word from it and show it to them for approval.
- **They don't have one:** interview them about their career, one role at a time, newest first:
  title, company, dates, what they owned, and 2-4 results *with numbers* (ask "how much / how
  many / how fast?" when they give an adjective). Then education, certifications, languages,
  tools. Build a clean two-page Word resume from the answers and ask them to check every line.
  Tell them plainly: a resume built from an interview is only as good as the answers.

### Step 3: facts and voice
Write `profile/FACTS.md` (facts only, nothing invented, every number as they said it) and
`profile/STYLE.md` (tone: formal or conversational, words they like or dislike, anything their
current resume already does well). Ask about anything that limits where they can work
(work authorization, relocation, languages) and record it as stated. Never add it to documents
unless they ask.

### Step 4: what they are looking for -> `profile/SETTINGS.json`
Cover, in this order, and write each answer into SETTINGS.json as you go:
1. **Country.**
2. **Tracks.** Their main target (Track A): the kind of role, the level, and 5-15 job titles
   people actually use for it (suggest common variants they did not name, and let them choose).
   Then ask if they want a second, fallback track (Track B), e.g. a step down in level or an
   adjacent field. One track is fine.
3. **Locations,** in priority order, each with a radius in km, plus whether remote counts and
   how much they want it.
4. **Title filter.** Suggest 10-20 words that mark titles clearly outside their field
   (from their tracks) and let them edit the list.
5. **Documents.** Any rule for every resume and cover letter ("always mention X",
   "never use the word Y"). Store as `documents.standing_notes`.

### Step 5: how a posting is scored -> `profile/SCORING.md` + `scoring.weights`
Explain the default in one sentence: **Role fit 40, Location 25, Seniority 20, Type & pay 15**,
with role fit weighted highest because a high location weight tends to push nearby jobs that are
not a real fit above 60. Ask whether that matches how they decide, and adjust (weights must sum to
100). Then ask: "What is an automatic no for you, however good the rest looks?" and write those as
hard red flags in SCORING.md. Read the final rubric back to them.

### Step 6: where to search -> `sources` (Full setup; in Basic, they paste postings)
Offer the built-in sources and turn on the ones they want:
- **Company career sites:** ask for 10-40 employers they would love to work for (suggest some from
  their field and region). Write them to `company_targets.csv` with the careers-page URL; find the
  URL by searching, and show the list for approval.
- **Government of Canada Job Bank** (Canada only): which nights (default Monday and Thursday).
- **Job-alert emails in Gmail** (Gmail only for now): help them set up alerts on job sites they
  already use, then agree on a Gmail search that finds those emails (e.g. `from:` the alert
  sender). Explain: alert emails carry only a title and a link, so the nightly description step
  fills in the full posting before scoring.
- **Anything else:** see section 6, "Add a search source".

### Step 7: close the setup
Summarise in a short list what was set up, what was left out, and the three things they can say
next: "score this posting", "prepare my application for <job>", "prepare me for my interview at
<company>". For Full setup, continue with the installation steps in the guide (scan, dashboard,
scheduled tasks).

### Changing settings later
The user never edits files. "Widen the radius to 50 km", "add project manager titles",
"stop searching Hamilton", "give location less weight": update SETTINGS.json (and SCORING.md when
it is about scoring), read the change back in one line, and say when it takes effect (the next
scan or the next scoring).

## 2. Finding jobs

- **Basic:** the user pastes a posting or a link; or, when asked, search live with web search and
  page fetching (government job boards, company careers pages, applicant-tracking boards that
  open without a login). Add each posting with
  `python scripts/pipeline.py add-job --source "<site>" --title "<t>" --company "<c>" --location "<l>" --url "<u>" --description "<full text>"`,
  then score it (section 3). A search is done only when everything found is scored.
- **Full:** the nightly scan (`scripts/run_scan.py`, started by `scripts/nightly_scan.bat`) runs
  every enabled source in SETTINGS.json; the morning agent adds jobs from alert emails, scores
  everything new and refreshes the dashboard.
- Never sign in to a site, and never browse behind the user's personal login, unless they ask for
  that specific site in that moment. Sites the user wants scanned regularly go through section 6.

## 3. Scoring

Pull the unscored postings and score them against `profile/SCORING.md` — not a fixed formula, but
not a bare gut-feel either:
```bash
python scripts/pipeline.py list --status New
```
For each one, read the job description against `profile/FACTS.md` and apply the weighted criteria
(weights from `profile/SETTINGS.json` -> `scoring.weights`) and the hard red-flag rules written into
`profile/SCORING.md` during setup. Be honest — a 95
should mean "I would apply to this today," not "this is vaguely related." Write the score, track
tag, and a per-criterion breakdown back:
```bash
python scripts/pipeline.py score --url "<job url>" --score 87 --track A --note "Role 34/40, Location 25/25, Seniority 16/20, Type 12/15 = 87"
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
0. Run `python scripts/prepare_application.py context --url "<job url>"` when the job is in the
   pipeline: it returns the description, the right template and its exact section headings,
   FACTS.md, STYLE.md, the user's standing document rules (`standing_notes`, follow every one)
   and the page limit (`resume_max_pages`). `prepare_application.py build` then fills the
   templates and checks the page count. The manual steps below are the same flow by hand.
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
5. The Word file is the deliverable. Keep the resume within `documents.resume_max_pages` from
   SETTINGS.json; if it runs over, shorten the least relevant bullets (older roles first), never
   fonts or margins.
6. Before calling it done, re-read the draft against the JD and confirm the terms identified in
   step 2 actually made it in somewhere reasonable. If one didn't fit naturally, say so rather than
   forcing it in or skipping it silently.
7. Update the pipeline once the user actually applies (never before they confirm they did):
   `python scripts/pipeline.py status --url "<job url>" --status Applied` — this stamps
   `applied_date` and sets a follow-up date automatically. If the application went through a
   referral, record who: `python scripts/pipeline.py referral --url "<job url>" --name "<person>" --note "<how>"`.
8. Delete temp JSON files. Confirm final file paths/names to the user and present the files.

## 6. Add a search source

Trigger: "add <site> to my search", "also look at <url>", "watch <company>".

1. **A company's own careers page:** add a row to `company_targets.csv` (name, careers URL,
   location). Run `python scripts/scan_companies.py --only "<name>"` once to confirm postings are
   read; if the page needs a browser, say so and set `how` accordingly.
2. **Any other site (job board, association board, RSS feed):** before adding it, tell the user
   once, plainly: "You're responsible for checking that this site's terms of use allow automated
   reading. Many large job sites do not." Add it only after they confirm, as an entry in
   `sources.extra` with `name`, `url`, `how` (`rss`, `html_list` or `browser`),
   `added_on` and `terms_checked_by_user: true`. Write or adapt the small reader script it needs
   in `scripts/sources/`, run it once, and show the user what it found.
3. Read back in one line what was added and when it will first run.
4. Removing a source: set `enabled: false` or delete the entry; never delete jobs already found.

## 7. Interview preparation

Trigger: "prepare me for my interview at <company>", "let's practise", or a kit request from the
dashboard (the sync agent follows `agents/sync_agent.md`).

1. Context: `python scripts/prepare_application.py context --url "<job url>"` (or the posting the
   user pastes), plus the resume and cover letter that were actually sent if they exist.
2. Research with web search: what the company does, size, leadership, recent news with dates
   (last ~6 months), where the role sits, and what employees say. At least two independent sources;
   mark company-reported figures as such; never invent a number. Interviewers: public pages only.
3. Deliver: a short research brief tied to the posting; 12-15 likely questions, each with a
   spoken-style answer of 80-150 words built ONLY from FACTS.md at its exact scope (STAR where it
   fits) and a one-line tip; always include "tell me about yourself", "why this role", the hardest
   gap between the user and the posting (answered honestly, bridging to what they have), seniority
   fit, and the posting's own success criteria; 5-6 sharp questions to ask them; a thank-you email
   draft with [placeholders].
4. Practice mode: ask one question at a time, wait for the answer, then rate it 1-10 with two
   things that worked, two to improve, and a stronger version in the user's own facts.
5. After the interview: offer to log the round and a summary (dashboard "Interview happened" or
   `pipeline.py status ... --status Interviewing`) and draft the thank-you note. Never send it.

## 8. Full setup (Claude-led)

Follow GUIDE.md chapter 3. In short, with the user's computer folder connected:
1. [You] Python 3.10+ installed with "Add python.exe to PATH". [Claude] run `install.bat` in the
   project folder (creates .venv, installs the libraries and the scanners' browser, creates the files).
2. [Claude] Setup interview (section 1) including sources; build `company_targets.csv`.
3. [Claude] Test: `.venv\Scripts\python.exe scripts\run_scan.py` once, show what it found.
4. [Claude] Schedule the scan: `scripts\install_nightly_scan.bat <HH:MM>`; tell the user how to tick
   "Wake the computer to run this task" if they want it.
5. [You] Connect Gmail to Claude if they want job-alert emails or employer replies.
6. [Claude] Create the dashboard: copy `dashboard/job-search-hub.template.html`, fill
   `{{PROJECT_DIR_JS}}` (the project path with doubled backslashes and a trailing backslash),
   publish it as a private artifact with the capabilities db, sample, and mcp
   (server "Claude Code Remote", tool update_trigger).
7. [Claude] Create the two scheduled tasks from `agents/`: the sync task (no schedule; the page
   books it) and the morning task (weekdays, about an hour after the scan). Fill the placeholders,
   then put the sync task's id into the dashboard's `{{SYNC_TRIGGER_ID}}` and republish. [You]
   approve each task when asked.
8. [Claude] First run: press Refresh on the dashboard together, check the board fills in, and
   explain the home page in two sentences.

## Privacy

Everything under `profile/`, `templates/`, `pipeline/` and `output/` is the user's personal data.
It stays in their own files and their own Claude account; this project has no server and sends
nothing anywhere except the job pages Claude reads while searching. The dashboard is a private
artifact that only the user can open unless they share it.
