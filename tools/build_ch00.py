"""build_ch00.py — rebuild decks/1851-Ch00.pptx for course 1851 revision A.2.

Idempotent: restores the carrier from decks/_bak0/1851-Ch00.pptx, then
  * keeps the title slide untouched,
  * rewrites the Course Objectives bullets in place (run-level edit, so
    paragraph/run formatting is preserved),
  * rewrites the Course Contents chapter rows in place (same technique),
  * appends a NEW slide "What's New in This Revision" with speaker notes.

Run (from 1851a2-author-input/):
    /Users/iahmad/Creator/Courses_and_conferences/LT/.venv-courseware/bin/python tools/build_ch00.py
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

SRC = ROOT / "decks" / "_bak0" / "1851-Ch00.pptx"
OUT = ROOT / "decks" / "1851-Ch00.pptx"

# --- A.2 content (verbatim) ---------------------------------------------------
OBJECTIVES = [
    "Describe how predictive and generative AI apply across the software development lifecycle",
    "Apply machine-learning essentials and explain how they underpin modern LLMs",
    "Engineer prompts and integrate the OpenAI API into software engineering tasks",
    "Use generative AI in requirements, design, coding, testing (TDD), and deployment workflows",
    "Design and run evaluation suites (evals) for LLM-powered features and agents",
    "Build AI agents and agentic workflows with tool use and human oversight",
    "Apply quality, ethics, and security practices to AI-enabled software",
]

CHAPTERS = [
    "Chapter 1: Introduction to AI in Software Development",
    "Chapter 2: Quality Characteristics and Ethics in AI",
    "Chapter 3: Machine Learning Essentials",
    "Chapter 4: Prompt Engineering with OpenAI",
    "Chapter 5: Generative AI Across the Software Development Lifecycle",
    "Chapter 6: Evaluating Generative AI Systems",
    "Chapter 7: AI Agents and Agentic Workflows",
    "Chapter 8: AI Security and Vulnerability Testing",
    "Chapter 9: Course Summary",
]

WHATS_NEW = [
    "GenAI-first approach, built around the OpenAI API",
    "New chapters on AI agents, agentic workflows, and evals",
    "Generative AI applied across the SDLC, including test-driven development (TDD)",
    "Machine learning condensed into a single essentials chapter",
    "New hands-on labs throughout the course",
]
WHATS_NEW_NOTES = (
    "This revision repositions the course around generative AI, with the OpenAI API "
    "as the primary hands-on platform. The two biggest additions are a chapter on AI "
    "agents and agentic workflows and a dedicated chapter on evaluating generative AI "
    "systems (evals). Traditional machine learning is condensed into one essentials "
    "chapter, and generative AI is now applied across every phase of the SDLC, "
    "including test-driven development. New hands-on labs accompany the updated chapters."
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


def rewrite_contents(slide, chapters) -> None:
    """Rewrite the chapter rows of the Course Contents TOC in place, keeping the
    leading 'Introduction and Overview' row and the trailing 'Next Steps' row.
    Rows keep the template's 'Chapter N<TAB>Title' convention."""
    tf = pt._body_placeholder(slide).text_frame
    paras = tf.paragraphs
    assert len(paras) == len(chapters) + 2, f"expected {len(chapters)+2} TOC rows, got {len(paras)}"
    for i, ch in enumerate(chapters, start=1):
        label, title = ch.split(": ", 1)
        set_para_text(paras[i], f"{label}\t{title}")


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
    assert len(prs.slides) == 3, f"carrier should have 3 slides, has {len(prs.slides)}"

    replace_bullets(prs.slides[1], OBJECTIVES)          # slide 1: Course Objectives
    rewrite_contents(prs.slides[2], CHAPTERS)           # slide 2: Course Contents
    pt.append_content(prs, "What's New in This Revision", WHATS_NEW,
                      notes=WHATS_NEW_NOTES)            # slide 3: NEW, after contents

    sk.save(prs, str(OUT))
    validate(OUT, expect_slides=4)


if __name__ == "__main__":
    main()
