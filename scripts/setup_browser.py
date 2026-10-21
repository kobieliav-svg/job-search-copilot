#!/usr/bin/env python3
"""Shared Playwright persistent-context helper used by login.py and every scanner.

Each site gets its own on-disk browser profile under .browser_profile/<site>/, so a
one-time manual login (see login.py) is reused by every future headless scan without
re-authenticating. This automates the user's own logged-in session -- it does not
store or replay credentials, and it is meant for personal, human-paced use only.
"""
import os
import sys
import threading
from contextlib import contextmanager
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
PROFILE_ROOT = ROOT / ".browser_profile"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Default wall-clock budget for one scan() call (one query/location pair).
# Generous for normal operation (page load + up to ~40 gently-throttled card
# reads), well below what a genuinely stuck run would take.
DEFAULT_SCAN_TIMEOUT = 90


def run_with_timeout(func, *args, timeout_seconds=DEFAULT_SCAN_TIMEOUT, **kwargs):
    """Run `func` and hard-exit the process if it hasn't returned within
    timeout_seconds.

    Playwright calls (browser launch, page.goto, element waits) each have
    their own internal timeouts, but a stuck OS-level condition -- most
    commonly a stale lock on the persistent profile directory when a new
    Chromium is launched immediately after a previous one against the same
    .browser_profile/<site>/ closes -- can hang past all of them. Rather than
    add narrower timeouts at each individual Playwright call (and still miss
    whatever we didn't think of), this is a last-resort watchdog around the
    whole scan: if it fires, something is genuinely stuck, not just slow.

    Deliberately hard-exits (os._exit) rather than raising: a stuck
    Playwright/Chromium call cannot reliably be interrupted from a Python
    thread, so this guarantees the process actually terminates -- with a
    non-zero exit code and a clear stderr message -- instead of leaving a
    scan silently hung (or silently "succeeding" with 0 results) in a
    pipeline run or CI-style batch of scans."""
    timer = threading.Timer(timeout_seconds, _timeout_kill, args=(timeout_seconds,))
    timer.daemon = True
    timer.start()
    try:
        return func(*args, **kwargs)
    finally:
        timer.cancel()


def _timeout_kill(timeout_seconds):
    print(
        f"TIMEOUT: scan did not finish within {timeout_seconds}s -- likely a stuck "
        "browser launch or page load (e.g. a stale lock on the persistent profile "
        "directory from a rapid successive launch, or an auth checkpoint/CAPTCHA). "
        "Failing fast instead of hanging silently. If this keeps happening for one "
        "site, check .browser_profile/<site>/ for a stray lock file, or re-run "
        "scripts/login.py <site> in case the session expired.",
        file=sys.stderr,
    )
    # os._exit() skips Python's normal shutdown, which is the whole point (a
    # stuck Playwright call can't be interrupted, so a clean exit isn't an
    # option) -- but that also skips the final stdout/stderr flush, silently
    # dropping any buffered print() output from earlier in this run (e.g. a
    # prior successful "added N job(s)" line) along with this message.
    # Flush explicitly first so nothing already-completed gets lost.
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(1)


@contextmanager
def browser_context(site: str, headless: bool = True):
    """Yield a Playwright BrowserContext backed by a persistent profile for `site`."""
    profile_dir = PROFILE_ROOT / site
    profile_dir.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir),
            headless=headless,
            user_agent=USER_AGENT,
            viewport={"width": 1400, "height": 1000},
        )
        try:
            yield context
        finally:
            context.close()
