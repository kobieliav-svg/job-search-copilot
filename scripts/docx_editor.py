"""In-place .docx section editing.

The whole point: take the user's *real* resume file and edit the text inside it,
rather than generating a new document from scratch. This preserves the user's actual
fonts, margins, spacing, and layout -- the formatting fidelity a from-scratch template
can never quite match, and the thing most likely to break (especially for non-Latin
scripts) when a document is rebuilt programmatically from zero.

Approach: find a heading paragraph by its text, then treat every paragraph after it
(up to the next heading-styled paragraph) as that section's content. To swap in new
bullet text, clone the XML of the first bullet paragraph in the section (so its
formatting -- font, bullet character, indentation -- carries over exactly), then set
new text on each clone and drop the original bullets.
"""
import copy

from docx import Document
from docx.text.paragraph import Paragraph


def _is_heading(paragraph: Paragraph) -> bool:
    style_name = (paragraph.style.name or "").lower()
    return "heading" in style_name or "title" in style_name


# Styles used for actual bullet/body content in resumes built by hand in Word
# (no real Heading styles applied). Anything else non-empty is treated as the
# start of the next job/section -- see _looks_like_section_break below.
_BODY_STYLES = {"list bullet", "list paragraph"}


def _looks_like_section_break(paragraph: Paragraph) -> bool:
    """Heuristic fallback for resumes with no real Word heading styles: a
    non-bullet ("Normal"-style) paragraph with text is almost always the next
    job title or section header, not more body content for the current one.
    Without this, section_paragraphs() below has no way to find where a
    section ends and silently swallows everything to the end of the document
    (deleting it) when a bullet section shorter than that gets written back.

    Some job-title lines are themselves typed with "List Bullet"/"List
    Paragraph" style (matching the bullets under them), so the style check
    alone misses them. A second signal catches those: a job-title line is
    fully bold ("Title | Company | Year" all bold), while real bullets only
    use bold for emphasis on part of the sentence -- verified with zero false
    positives/negatives across every paragraph in the Professional Experience
    section of resume_master_senior.docx."""
    if not paragraph.text.strip():
        return False
    style_name = (paragraph.style.name or "").lower()
    if style_name not in _BODY_STYLES:
        return True
    runs_with_text = [r for r in paragraph.runs if r.text.strip()]
    return bool(runs_with_text) and all(r.bold for r in runs_with_text)


def find_heading_index(doc: Document, heading_text: str):
    target = heading_text.strip().lower()
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip().lower() == target:
            return i
    # fallback: startswith match, in case the heading has extra trailing text
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip().lower().startswith(target):
            return i
    return None


def section_paragraphs(doc: Document, heading_text: str):
    """Return (heading_index, [paragraph indices belonging to this section])."""
    idx = find_heading_index(doc, heading_text)
    if idx is None:
        return None, []
    # The "looks like a new section" guess is only for resumes typed without Word heading styles.
    # When the resume uses real Heading styles, a section simply ends at the next heading.
    styled = any(_is_heading(p) and p.text.strip() for p in doc.paragraphs)
    body = []
    for j in range(idx + 1, len(doc.paragraphs)):
        para = doc.paragraphs[j]
        if _is_heading(para) and para.text.strip():
            break
        if not styled and _looks_like_section_break(para):
            break
        body.append(j)
    return idx, body


def _replace_inline_section(doc: Document, heading_text: str, text: str) -> bool:
    """Fallback for headings hand-typed as a line inside a larger paragraph
    (manual Shift+Enter line breaks) rather than their own paragraph -- e.g. a
    contact-info line, "PROFESSIONAL SUMMARY", and the summary text itself all
    sharing one <w:p>. find_heading_index/section_paragraphs can't address
    that, so scan raw paragraph text for the heading and replace everything
    after it (up to the paragraph's end) at the run level, preserving the
    formatting of whichever run the old body text lived in.

    Offsets are computed against the concatenation of paragraph.runs only,
    NOT paragraph.text -- a paragraph containing a hyperlink (e.g. a LinkedIn
    URL) has hyperlink runs nested under a separate <w:hyperlink> element that
    .text includes but .runs does not, so the two strings can differ in
    length and any offset found in one is meaningless against the other.
    """
    target = heading_text.strip().lower()
    for para in doc.paragraphs:
        full = "".join(run.text for run in para.runs)
        pos = full.lower().find(target)
        if pos == -1:
            continue
        body_start = pos + len(heading_text)
        if body_start < len(full) and full[body_start] == "\n":
            body_start += 1
        if body_start >= len(full):
            continue  # heading text but nothing after it -- not this case

        cum = 0
        cut_run_idx = None
        cut_offset = None
        for ri, run in enumerate(para.runs):
            run_len = len(run.text)
            if cum + run_len >= body_start:
                cut_run_idx = ri
                cut_offset = body_start - cum
                break
            cum += run_len
        if cut_run_idx is None:
            continue

        cut_run = para.runs[cut_run_idx]
        cut_run.text = cut_run.text[:cut_offset]

        remaining = para.runs[cut_run_idx + 1:]
        if remaining:
            remaining[0].text = text
            for extra in remaining[1:]:
                extra.text = ""
        else:
            new_run = para.add_run(text)
            new_run.bold = False
            new_run.font.name = cut_run.font.name
            new_run.font.size = cut_run.font.size
        return True
    return False


def _set_text_preserve_format(paragraph: Paragraph, text: str) -> None:
    """Replace a paragraph's visible text while keeping the first run's formatting."""
    if not paragraph.runs:
        paragraph.add_run(text)
        return
    first = paragraph.runs[0]
    first.text = text
    for extra in paragraph.runs[1:]:
        extra.text = ""


def replace_bullets(doc: Document, heading_text: str, new_bullets: list[str]) -> bool:
    """Replace the bullet paragraphs under `heading_text` with `new_bullets`.

    Clones the XML of the section's first paragraph as the formatting template for
    every new bullet, so bullet style/indentation/font is preserved exactly.
    Returns True if the section was found and edited.
    """
    heading_idx, body_idx = section_paragraphs(doc, heading_text)
    if heading_idx is None or not body_idx:
        return False

    template_para = doc.paragraphs[body_idx[0]]
    template_el = template_para._p

    # Set text on however many of the existing paragraphs we can reuse.
    reusable = body_idx[: len(new_bullets)]
    for i, text in zip(reusable, new_bullets):
        _set_text_preserve_format(doc.paragraphs[i], text)

    # Remove leftover paragraphs if we have fewer new bullets than before.
    if len(new_bullets) < len(body_idx):
        for i in reversed(body_idx[len(new_bullets):]):
            p = doc.paragraphs[i]._p
            p.getparent().remove(p)

    # Add paragraphs (cloned from the template) if we have more new bullets than before.
    if len(new_bullets) > len(body_idx):
        anchor = doc.paragraphs[body_idx[-1]]._p
        for text in new_bullets[len(body_idx):]:
            clone = copy.deepcopy(template_el)
            anchor.addnext(clone)
            anchor = clone
            new_para = Paragraph(clone, template_para._parent)
            _set_text_preserve_format(new_para, text)

    return True


def replace_paragraph_text(doc: Document, heading_text: str, text: str) -> bool:
    """Replace the single paragraph immediately following `heading_text` (e.g. a
    one-line professional summary) with `text`, preserving its formatting."""
    # The `body_idx` paragraphs found here can be pure empty spacer paragraphs when
    # heading_text and its real body text live in the same <w:p> (a hand-typed heading
    # followed by Shift+Enter, not its own paragraph) -- section_paragraphs() has no way
    # to see that from paragraph boundaries alone. Writing into an empty spacer there would
    # silently leave the real (stale) text untouched, so require at least one non-empty
    # body paragraph before taking this path; otherwise fall through to the inline handler.
    heading_idx, body_idx = section_paragraphs(doc, heading_text)
    if heading_idx is not None and body_idx and any(doc.paragraphs[i].text.strip() for i in body_idx):
        _set_text_preserve_format(doc.paragraphs[body_idx[0]], text)
        # collapse any additional paragraphs in that section into the first one
        for i in reversed(body_idx[1:]):
            p = doc.paragraphs[i]._p
            p.getparent().remove(p)
        return True
    return _replace_inline_section(doc, heading_text, text)


def replace_placeholder(doc: Document, placeholder: str, value: str) -> int:
    """Simple find/replace for inline placeholders like {{COMPANY}}, run-safe.

    Handles the common case where python-docx splits a placeholder across multiple
    runs by rebuilding the paragraph's run list only when the placeholder is found
    in the paragraph's full text.
    """
    count = 0
    for para in doc.paragraphs:
        if placeholder in para.text:
            full = para.text.replace(placeholder, value)
            _set_text_preserve_format(para, full)
            count += 1
    return count
