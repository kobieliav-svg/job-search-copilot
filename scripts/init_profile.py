#!/usr/bin/env python3
"""Create the user's files from the blank templates: profile/FACTS.md, STYLE.md, SCORING.md,
SETTINGS.json, company_targets.csv and pipeline/jobs.xlsx.

Safe to run any number of times: an existing file is never overwritten, so re-running setup
never clobbers a profile that is already filled in.
"""
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PROFILE = ROOT / "profile"

PAIRS = [
    ("FACTS.template.md", "FACTS.md"),
    ("STYLE.template.md", "STYLE.md"),
    ("SCORING.template.md", "SCORING.md"),
    ("SETTINGS.template.json", "SETTINGS.json"),
]


def main() -> int:
    PROFILE.mkdir(exist_ok=True)
    created = []
    skipped = []
    for template_name, target_name in PAIRS:
        template = PROFILE / template_name
        target = PROFILE / target_name
        if not template.exists():
            print(f"ERROR: missing template {template}", file=sys.stderr)
            return 1
        if target.exists():
            skipped.append(target_name)
            continue
        shutil.copy(template, target)
        created.append(target_name)

    csv_t, csv = ROOT / "company_targets.template.csv", ROOT / "company_targets.csv"
    if csv_t.exists() and not csv.exists():
        shutil.copy(csv_t, csv)
        created.append("company_targets.csv")
    if not (ROOT / "pipeline" / "jobs.xlsx").exists():
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import pipeline
        pipeline.init()
        created.append("pipeline/jobs.xlsx")

    if created:
        print("Created: " + ", ".join(created))
    if skipped:
        print("Already existed, left untouched: " + ", ".join(skipped))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
