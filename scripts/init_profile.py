#!/usr/bin/env python3
"""Scaffold profile/FACTS.md, profile/STYLE.md, and profile/SCORING.md from their templates.

Safe to run multiple times: never overwrites an existing file, so re-running onboarding
doesn't clobber a profile that's already been filled in.
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

    if created:
        print("Created: " + ", ".join(created))
    if skipped:
        print("Already existed, left untouched: " + ", ".join(skipped))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
