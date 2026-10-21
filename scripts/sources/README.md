# Extra search sources

Each file here reads one extra job source the user asked Claude to add
("add <site> to my search"). It is listed in `profile/SETTINGS.json` -> `sources.extra` with
`"script": "<file name>"`, and `run_scan.py` runs it every scan night.

A source script only has to: fetch the postings, build a list of
`{"source", "title", "company", "location", "url"}` and call `pipeline.add_jobs(items)`.
The description step and the scoring then work exactly as for every other source.

Before adding a site, the user is responsible for checking that its terms of use allow
automated reading. Many large job sites do not. Claude says this once before adding a source,
and records `"terms_checked_by_user": true` in SETTINGS.json.

`rss_example.py` is a working template for any site that publishes an RSS feed of its jobs.
