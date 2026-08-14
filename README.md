# job-search-copilot

A free [Claude](https://claude.ai) skill that scores job postings against your real profile,
tracks everything in one spreadsheet, and drafts tailored resumes and cover letters — inside a
chat, no install, no terminal, no coding.

**It doesn't invent facts.** One file (`profile/FACTS.md`) holds everything true about you.
A second file (`profile/STYLE.md`) holds how you write it up. A third (`profile/SCORING.md`) holds
your own scoring rubric. Every generated document and every score is built only from those.

## Install — one step

1. Download [`job-search-copilot.skill`](./job-search-copilot.skill) from this repo.
2. In [Claude, Cowork mode](https://claude.ai), drag the file in — a **Save skill** button
   appears. Click it.
3. Say **"set up my job search."** That's it.

No Python, no virtual environment, no API key beyond the Claude session you're already in.

## What it does

- Scores every posting 0–100 against a profile built from your real resume plus a short interview
  — role fit, location, seniority, and whatever else matters most to you, weighted the way you
  actually think about it.
- Tracks everything in one Excel workbook Claude creates and updates for you.
- Drafts tailored resumes and cover letters, filled into your own real resume file — never built
  from scratch, never inventing experience you don't have.
- Tracks follow-ups automatically on applications and outreach messages.
- Never sends or submits anything without asking you first.

See [`SKILL.md`](./SKILL.md) for the full, exact instructions Claude follows — it's plain
markdown, worth a skim if you want to know precisely what it will and won't do.

## Repo layout

```
SKILL.md                    # what Claude reads to run this skill
job-search-copilot.skill    # the installable package (zip of everything below + SKILL.md)
profile/
  FACTS.template.md         # blank template — your filled FACTS.md stays in your own session
  STYLE.template.md         # blank template
  SCORING.template.md       # blank template
scripts/                    # pipeline tracking + resume/cover-letter generation (pure Python,
                             # runs in Claude's built-in sandbox, no install needed)
```

## Privacy

This has no server and no account of its own. Your resume, job list, and messages live only in
your own Claude session and files — nothing is sent anywhere except the job pages Claude browses
on your behalf while searching, at your request.

## Good to know

- This is a research and drafting assistant, not an auto-apply bot. You review and send everything
  yourself.
- Job "scanning" happens live, in conversation — Claude searches and browses on request rather than
  running a scheduled background scraper. Great for focused sessions; not built for unattended
  24/7 monitoring.
- Sites that require your personal login (most of LinkedIn, for example) can only be browsed if
  you're already logged in on your own connected browser, and only when you explicitly ask.

## License

MIT — see [LICENSE](./LICENSE).
