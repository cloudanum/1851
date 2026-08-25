"""build_ch09.py — rebuild decks/1851-Ch09.pptx for course 1851 revision A.2.

Idempotent: restores the carrier from decks/_bak0/1851-Ch09.pptx, then
  * keeps the "Course Summary" chapter-divider slide untouched,
  * rewrites the objectives-recap slide with the 7 A.2 objectives in place
    (run-level edit, so paragraph/run formatting is preserved),
  * inserts a NEW "Course Summary" content slide (6 takeaway bullets, with
    speaker notes) between the divider and the recap.

Run (from 1851a2-author-input/):
    /Users/iahmad/Creator/Courses_and_conferences/LT/.venv-courseware/bin/python tools/build_ch09.py
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import slidekit as sk      # noqa: E402  (import also patches safe slide partnames)
import pptx_tools as pt    # noqa: E402

SRC = ROOT / "decks" / "_bak0" / "1851-Ch09.pptx"
OUT = ROOT / "decks" / "1851-Ch09.pptx"

# --- A.2 content ---------------------------------------------------------------
OBJECTIVES = [
    "Describe how predictive and generative AI apply across the software development lifecycle",
    "Apply machine-learning essentials and explain how they underpin modern LLMs",
    "Engineer prompts and integrate the OpenAI API into software engineering tasks",
    "Use generative AI in requirements, design, coding, testing (TDD), and deployment workflows",
    "Design and run evaluation suites (evals) for LLM-powered features and agents",
    "Build AI agents and agentic workflows with tool use and human oversight",
    "Apply quality, ethics, and security practices to AI-enabled software",
]

TAKEAWAYS = [
    "Generative AI is an engineering discipline — practiced through prompts, APIs, and evals",
    "The OpenAI API gives repeatable integration patterns for AI-enabled software",
    "TDD and evals act as quality gates for AI features",
    "AI agents and agentic workflows extend automation, with human oversight",
    "Security and ethics apply throughout the SDLC, not as an afterthought",
    "AI evolves quickly — continuous learning is part of the job",
]
SUMMARY_NOTES = (
    "Generative AI pays off when it is treated as an engineering discipline: deliberate "
    "prompting, clean OpenAI API integration, and automated evaluation. Test-driven "
    "development and evals provide the quality gates for AI features, while agents and "
    "agentic workflows extend automation under human oversight. Security and ethics stay "
    "cross-cutting concerns across the whole SDLC. The tooling will keep changing, so "
    "encourage participants to keep experimenting after the course."
)


# --- run-preserving text edits -------------------------------------------------
def set_para_text(p, text: str) -> None:
    """Set a paragraph's text on its first run and drop any extra runs, so the
    surviving run keeps its formatting. Paragraph properties are untouched."""
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p.add_run().text = text


def replace_bullets(slide, lines) -> None:
    """Rewrite the main body placeholder to exactly `lines`, editing paragraphs
    in place. Surplus old bullets are removed; the template's trailing empty
    paragraph is kept."""
    tf = pt._body_placeholder(slide).text_frame
    paras = tf.paragraphs
    assert len(paras) >= len(lines) + 1, f"expected >= {len(lines)+1} paragraphs, got {len(paras)}"
    assert not paras[-1].text.strip(), "expected a trailing empty paragraph"
    for i, line in enumerate(lines):
        set_para_text(paras[i], line)
    for p in paras[len(lines):-1]:
        p._p.getparent().remove(p._p)


# --- validation ----------------------------------------------------------------
def validate(path: Path, expect_slides: int) -> int:
    prs = sk.open_prs(path)
    n = len(prs.slides)
    assert n == expect_slides, f"{n} slides != expected {expect_slides}"
    print(f"\n{path}: {n} slides")
    for i, s in enumerate(prs.slides):
        title = sk.slide_title(s)
        print(f"  [{i}] {title!r}")
        assert title.strip(), f"slide {i} has an empty title"
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    dupes = sorted({nm for nm in names if names.count(nm) > 1})
    assert not dupes, f"duplicate zip partnames: {dupes}"
    r = subprocess.run(["unzip", "-t", str(path)], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    print("  zip partnames unique:", len(names), "parts;", r.stdout.strip().splitlines()[-1])
    return n


def main() -> None:
    shutil.copyfile(SRC, OUT)  # idempotent: always start from the pristine carrier
    prs = sk.open_prs(OUT)
    assert len(prs.slides) == 2, f"carrier should have 2 slides, has {len(prs.slides)}"

    # slide 1 (recap): rewrite in place BEFORE inserting, so its index is stable
    replace_bullets(prs.slides[1], OBJECTIVES)

    # NEW summary content slide: append, then move between divider and recap
    pt.append_content(prs, "Course Summary", TAKEAWAYS, notes=SUMMARY_NOTES)
    sk.move_slide(prs, len(prs.slides._sldIdLst) - 1, 1)

    sk.save(prs, str(OUT))
    validate(OUT, expect_slides=3)


if __name__ == "__main__":
    main()
