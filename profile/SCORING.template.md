# Scoring rubric

Filled in by Claude during onboarding, together with the user, from `profile/FACTS.md` and a short
conversation about what "a great match" actually means for this person. Not a fixed formula —
but not a bare gut-feel either. Once written, every posting gets one 0-100 score plus a `track` tag,
scored against exactly this rubric so scores stay comparable to each other over time.

- **Track A** = <the user's primary target: e.g. "senior-level roles in field X">
- **Track B** = <the user's secondary/backup target, if they want one: e.g. "a step down in
  seniority" or "an adjacent field" — optional; skip Track B entirely if the user only wants one
  track, and drop `--track` from scoring commands>

```bash
python scripts/pipeline.py score --url "<job url>" --score 87 --track A --note "..."
```

## Weighted criteria

The weights live in `profile/SETTINGS.json` (`scoring.weights`) and must sum to 100. The default is
below. Role fit gets the most weight on purpose: when location weighs more, jobs that are nearby but
not a real fit tend to land above 60. Change the weights with the user during setup if they decide
differently, and keep this table in step with SETTINGS.json.

| Criterion | Default weight | Full marks means |
|---|---|---|
| Role fit | 40 | <the real duties (read the description, not just the title) match the user's actual target work, per FACTS.md -- check the *domain*, not just keyword overlap> |
| Location | 25 | <the user's top-priority locations/remote preference from SETTINGS.json get full marks; lower-priority ones get high but not full; anything else scores down> |
| Seniority match | 20 | <level matches the user's actual experience -- not a meaningful downgrade, not out of reach> |
| Type and pay | 15 | <full-time/contract preference from FACTS.md; deduct only if listed pay is clearly below a reasonable range for the level> |

Write the per-criterion breakdown into `--note` (not just a one-line verdict) so every score is
auditable later, e.g. `"Role 34/40, Location 25/25, Seniority 16/20, Type 12/15 = 87"`.

## Hard red flags (override the weighted score, don't just subtract points)

Ask the user directly: "Is there anything that's an automatic no for you, no matter how good the
rest of the posting looks?" Common examples to prompt with if they're not sure — keep only what's
actually true for them, delete the rest, and add anything specific they mention:

- **A language/certification/eligibility requirement they don't meet** -> score 0-10, note why.
- **A skills gap stated as required** (e.g. a technical degree or hands-on experience the user
  genuinely doesn't have) -> Role Fit scores low regardless of everything else, this is a real gap
  not a stylistic mismatch.
- **A seniority level clearly above or below their real ceiling/floor** -> cap Seniority low, don't
  give it a pass just because the rest of the posting reads well.
- **A specific "required"/"non-negotiable" line names something absent from FACTS.md** -> cap the
  score at 55 at most (so it does not reach the board) and name the gap in the note, but check
  FACTS.md carefully first — don't assume a gap that isn't actually there.
- **Templated/generic posting with pay mismatched to the role** -> add "CAUTION: possible low-
  quality/scam listing" to the note.
- **Title doesn't match the actual described duties** -> score the real duties, not the title, and
  say so in the note.
- **Posting text itself says the role is filled/closed** -> score it for the record, then archive
  it immediately with a note, instead of leaving it cluttering the active lists:
  `python scripts/pipeline.py status --url "<job url>" --status Archived --note "Closed per posting text"`.

## No hard cutoff (recommended default)

Don't filter the pipeline by a minimum score at first — score everything honestly and let the
scores themselves tell the story once there's a range to compare across. A cutoff can be added
later once the user has enough scored postings to sanity-check where the real bar sits.
