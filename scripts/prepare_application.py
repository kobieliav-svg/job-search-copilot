#!/usr/bin/env python3
"""Helper for "prepare my application" (the dashboard's Prepare button and chat requests).

Wraps the existing generate_resume.py / generate_cover_letter.py flow (SKILL.md section 5)
so the dashboard agent only has to write the tailored content. Never edits jobs.xlsx or
the templates; writes only inside output/applications/<folder>/.

1) context: everything needed to tailor documents for one job, as JSON on stdout.
     python scripts/prepare_application.py context --url "<job url>"
   Prints: job fields, the full description, which resume template to use (by track), the
   template's section headings (exact keys for bullet_replacements), FACTS.md, STYLE.md, the
   standing document rules from SETTINGS.json, and a new, unused output folder.

2) build: fill the templates, export PDFs, and write the page document.
     python scripts/prepare_application.py build --url "<job url>" --folder "<folder from context>" \
        [--resume resume.json] [--cover cover.json] [--meta meta.json]
   resume.json / cover.json use the generators' existing content format. meta.json is
   optional: {"ats_terms_used": [...], "ats_terms_missing": [...], "notes": "..."}.
   Word only, no PDF. The resume/cover letter .docx go straight into output/
   (a newer preparation for the same job replaces the older file of the same name); working files stay
   in output/_work/<folder>/.
   Writes <folder>/prep_doc.json (the document the agent uploads to the page) and prints
   {"ok": true, "prep_doc": "<path>"}.
"""
import argparse
import datetime
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path

import openpyxl
from docx import Document

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "pipeline" / "jobs.xlsx"
OUTPUT = ROOT / "output"
APPS = OUTPUT / "_work"  # working files per preparation (content JSON, meta, page document); the Word files go straight into output/
# Track A uses templates/resume_master.docx; track B uses templates/resume_master_B.docx if the user has
# a second resume layout for it, otherwise the same master.
TEMPLATES = {"A": ROOT / "templates" / "resume_master.docx", "B": ROOT / "templates" / "resume_master_B.docx"}
COVER_TEMPLATE = ROOT / "templates" / "cover_letter_master.docx"
MIN_DESC_CHARS = 200


def _s(v):
    return "" if v is None else str(v).strip()


def job_id(url):
    return hashlib.sha1(url.encode("utf-8")).hexdigest()[:16]


def find_job(url):
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb["Jobs"]
    rows = ws.iter_rows(values_only=True)
    hdr = [_s(h).lower() for h in next(rows)]
    ix = {n: i for i, n in enumerate(hdr)}
    for r in rows:
        if r and ix.get("url") is not None and _s(r[ix["url"]]) == url:
            return {n: _s(r[i]) if i < len(r) else "" for n, i in ix.items()}
    return None


def slug(text, n):
    words = re.sub(r"[^A-Za-z0-9 ]+", " ", text).split()
    return "".join(w[:1].upper() + w[1:] for w in words)[:n] or "Job"


def doc_text(path):
    d = Document(str(path))
    return "\n".join(p.text.strip() for p in d.paragraphs if p.text.strip())


def headings(path):
    """Paragraphs that look like section or job headings (what bullet/paragraph keys must match)."""
    out = []
    for p in Document(str(path)).paragraphs:
        for t in p.text.split("\n"):
            t = t.strip()
            if t and (t.isupper() or ("|" in t and len(t) < 140)) and "@" not in t:
                out.append(t)
    return out


def cmd_context(a):
    job = find_job(a.url)
    if not job:
        print(json.dumps({"ok": False, "error": "job_not_found"}))
        return 1
    track = (job.get("track") or "A").upper()[:1]
    track = track if track in TEMPLATES else "A"
    tpl = TEMPLATES[track]
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    folder = APPS / f"{slug(job.get('company', ''), 30)}_{slug(job.get('title', '').split(chr(10))[0], 40)}_{stamp}"
    desc = job.get("description", "")
    print(json.dumps({
        "ok": True,
        "job_id": job_id(a.url),
        "job": {k: job.get(k, "") for k in ("title", "company", "location", "score", "score_note", "status", "track", "url")},
        "description": desc,
        "has_description": len(desc) >= MIN_DESC_CHARS,
        "resume_template": str(tpl.relative_to(ROOT)),
        "resume_template_headings": headings(tpl),
        "cover_template": str(COVER_TEMPLATE.relative_to(ROOT)),
        "cover_placeholders": sorted(set(re.findall(r"\{\{[A-Z_0-9]+\}\}", doc_text(COVER_TEMPLATE)))),
        "facts": (ROOT / "profile" / "FACTS.md").read_text(encoding="utf-8"),
        "style": (ROOT / "profile" / "STYLE.md").read_text(encoding="utf-8"),
        "standing_notes": _doc_settings()["standing_notes"],
        "resume_max_pages": _doc_settings()["resume_max_pages"],
        "folder": str(folder.relative_to(ROOT)),
    }, ensure_ascii=False))
    return 0


def run(cmd):
    p = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, timeout=170)
    return p.returncode, (p.stdout + p.stderr).strip()


def _doc_settings():
    try:
        import settings
        return settings.documents()
    except SystemExit:
        return {"resume_max_pages": 2, "standing_notes": []}


def resume_layout(docx_path):
    """Counts the resume's pages through a temporary PDF made OUTSIDE the project folder (no PDF is
    ever kept). Returns the page count, or None if LibreOffice is not installed."""
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(docx_path)])
        pdf = Path(tmp) / (docx_path.stem + ".pdf")
        if not pdf.exists():
            return None
        try:
            from pypdf import PdfReader
            return len(PdfReader(str(pdf)).pages)
        except Exception:
            return None


def to_pdf(docx_path):
    code, out = run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)])
    pdf = docx_path.with_suffix(".pdf")
    return pdf if pdf.exists() else None, out


def cmd_build(a):
    job = find_job(a.url)
    if not job:
        print(json.dumps({"ok": False, "error": "job_not_found"}))
        return 1
    folder = (ROOT / a.folder).resolve()
    if APPS.resolve() not in folder.parents:
        print(json.dumps({"ok": False, "error": "folder_must_be_under_output/_work"}))
        return 1
    folder.mkdir(parents=True, exist_ok=True)
    track = (job.get("track") or "A").upper()[:1]
    tpl = TEMPLATES.get(track, TEMPLATES["A"])
    if not tpl.exists():
        tpl = TEMPLATES["A"]
    base = f"{slug(job.get('company', ''), 30)}_{slug(job.get('title', '').split(chr(10))[0], 40)}"
    warnings, files, doc = [], {}, {}

    for kind, content_arg, script, template in (
        ("resume", a.resume, "scripts/generate_resume.py", tpl),
        ("cover_letter", a.cover, "scripts/generate_cover_letter.py", COVER_TEMPLATE),
    ):
        if not content_arg:
            continue
        content_path = Path(content_arg)
        content = json.loads(content_path.read_text(encoding="utf-8"))
        saved = folder / f"{kind}_content.json"
        if content_path.resolve() != saved.resolve():
            saved.write_text(json.dumps(content, ensure_ascii=False, indent=2), encoding="utf-8")
        # one folder per job - output/applications/<base>/ until it is submitted;
        # scripts/file_documents.py moves it to output/Submitted/<base>/ once the job is Applied.
        (OUTPUT / "applications" / base).mkdir(parents=True, exist_ok=True)
        out_docx = OUTPUT / "applications" / base / f"{base}_{'Resume' if kind == 'resume' else 'CoverLetter'}.docx"
        # a new preparation for the same job replaces the older Word file
        # with the same name -- no dated duplicates piling up in output/.
        code, out = run([sys.executable, script, "--template", str(template), "--content", str(saved), "--output", str(out_docx)])
        warnings += [l.strip() for l in out.splitlines() if "WARNING" in l or "ERROR" in l]
        if code != 0 or not out_docx.exists():
            print(json.dumps({"ok": False, "error": f"{kind}_generation_failed", "detail": out[-600:]}))
            return 1
        if kind == "resume":
            pages = resume_layout(out_docx)
            files_pages = pages
            if pages is None:
                warnings.append("Could not count the resume's pages.")
            elif pages > _doc_settings()["resume_max_pages"]:
                warnings.append(f"PAGES: the resume is {pages} pages; the limit in SETTINGS.json is {_doc_settings()['resume_max_pages']}. Shorten it and build again.")
        text = doc_text(out_docx)
        if re.search(r"\{\{[A-Z_0-9]+\}\}", text):
            warnings.append(f"{kind}: a {{{{PLACEHOLDER}}}} was left unfilled")
        files[kind] = {"docx": str(out_docx.relative_to(ROOT))}
        if kind == "resume":
            files[kind]["pages"] = files_pages
        doc[kind] = {"text": text, "content": content}

    meta = json.loads(Path(a.meta).read_text(encoding="utf-8")) if a.meta else {}
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    prep = {
        "job_id": job_id(a.url), "url": a.url, "title": job.get("title", "").split("\n")[0],
        "company": job.get("company", ""), "state": "ready", "ready_at": now,
        "folder": str(folder.relative_to(ROOT)), "files": files, "warnings": warnings,
        "ats_terms_used": meta.get("ats_terms_used", []), "ats_terms_missing": meta.get("ats_terms_missing", []),
        "notes": meta.get("notes", ""),
    }
    prep.update(doc)
    out_path = folder / "prep_doc.json"
    out_path.write_text(json.dumps(prep, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"ok": True, "prep_doc": str(out_path.relative_to(ROOT)), "warnings": warnings}))
    return 0


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("context"); c.add_argument("--url", required=True)
    b = sub.add_parser("build"); b.add_argument("--url", required=True); b.add_argument("--folder", required=True)
    b.add_argument("--resume"); b.add_argument("--cover"); b.add_argument("--meta")
    a = ap.parse_args()
    return cmd_context(a) if a.cmd == "context" else cmd_build(a)


if __name__ == "__main__":
    raise SystemExit(main())
