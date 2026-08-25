"""build_engagement_ch00.py — engagement pass on decks/1851-Ch00.pptx (1851 A.2).

Idempotent: restores the deck from decks/_bak1/1851-Ch00.pptx (the post-build
A.2 snapshot), then
  * inserts a NEW icebreaker "Do Now: What Do You Want From This Course?"
    (Discussion layout) at index 2, right after Course Objectives,
  * rewrites the labs bullet of "What's New in This Revision" to
    "Nine hands-on labs — three per day — plus Do Now warm-ups in every chapter"
    (run-level edit, formatting preserved) and syncs its speaker notes.

Deck: 4 -> 5 slides.

Run (from 1851a2-author-input/):
    /Users/iahmad/Creator/Courses_and_conferences/LT/.venv-courseware/bin/python tools/build_engagement_ch00.py
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

SRC = ROOT / "decks" / "_bak1" / "1851-Ch00.pptx"
OUT = ROOT / "decks" / "1851-Ch00.pptx"

# --- NEW Do Now (Discussion layout) --------------------------------------------
DO_NOW = (
    "Do Now: What Do You Want From This Course?",
    ["Time box: 5 minutes — popcorn style, no raised hands needed",
     "Each person: ONE thing you want to walk out of this course able to do",
     "One sentence, starting \u2018I want to ...\u2019",
     "Your instructor maps each goal to the chapter that covers it",
     "Duplicates are welcome — they show what this room cares about"],
    "Run this right after the objectives, popcorn-style: one sentence, one goal each, "
    "no discussion yet. Capture goals on the whiteboard and map each to a chapter on the "
    "spot — it tells you which chapters to weight for this room. If a goal is out of scope, "
    "say so honestly and point to where they can get it. Keep it to 5 minutes; with a large "
    "group, pair-share first and hear only a few.",
)

# --- What's New bullet rewrite --------------------------------------------------
OLD_LABS_BULLET = "New hands-on labs throughout the course"
NEW_LABS_BULLET = "Nine hands-on labs — three per day — plus Do Now warm-ups in every chapter"
OLD_NOTES_LINE = "New hands-on labs accompany the updated chapters."
NEW_NOTES_LINE = ("The course now runs nine hands-on labs — three per day — and every "
                  "chapter opens with a short Do Now warm-up.")


# --- helpers ---------------------------------------------------------------------
def layout_named(prs, name):
    """Exact-name layout lookup (a substring match would grab the 'Numbered ...'
    Do Now variants first)."""
    for master in prs.slide_masters:
        for lay in master.slide_layouts:
            if (lay.name or "") == name:
                return lay
    raise KeyError(f"layout not found: {name}")


def add_engagement_slide(prs, layout_name, title, bullets, notes, at_index=None):
    """Add a slide on a named layout, fill title + main body placeholder, attach
    speaker notes, and optionally move it into position."""
    slide = prs.slides.add_slide(layout_named(prs, layout_name))
    sk.set_title(slide, title)
    body = pt._body_placeholder(slide)
    assert body is not None, f"no body placeholder on layout {layout_name!r}"
    tf = body.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = b
    sk.set_notes(slide, notes)
    if at_index is not None:
        sk.move_slide(prs, len(prs.slides._sldIdLst) - 1, at_index)
    return slide


def set_para_text(p, text: str) -> None:
    """Set a paragraph's text on its first run and drop any extra runs, so the
    surviving run keeps its formatting. Paragraph properties are untouched."""
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p.add_run().text = text


# --- validation ------------------------------------------------------------------
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
    shutil.copyfile(SRC, OUT)  # idempotent: always start from the _bak1 snapshot
    prs = sk.open_prs(OUT)
    assert len(prs.slides) == 4, f"expected 4 slides in _bak1 deck, has {len(prs.slides)}"

    # 1. NEW icebreaker after Course Objectives (idx 1) -> lands at idx 2
    add_engagement_slide(prs, "Discussion", DO_NOW[0], DO_NOW[1], DO_NOW[2], at_index=2)

    # 2. rewrite the labs bullet + notes on "What's New in This Revision"
    wn = next(s for s in prs.slides
              if sk.slide_title(s) == "What's New in This Revision")
    body = pt._body_placeholder(wn)
    hits = 0
    for p in body.text_frame.paragraphs:
        if p.text.strip() == OLD_LABS_BULLET:
            set_para_text(p, NEW_LABS_BULLET)
            hits += 1
    assert hits == 1, f"labs bullet not found exactly once (hits={hits})"
    nt = wn.notes_slide.notes_text_frame
    assert OLD_NOTES_LINE in nt.text, "What's New notes line not found"
    nt.text = nt.text.replace(OLD_NOTES_LINE, NEW_NOTES_LINE)

    sk.save(prs, str(OUT))
    validate(OUT, expect_slides=5)

    # spot-check the edit landed
    prs2 = sk.open_prs(OUT)
    wn2 = next(s for s in prs2.slides
               if sk.slide_title(s) == "What's New in This Revision")
    assert any(NEW_LABS_BULLET == p.text.strip()
               for p in pt._body_placeholder(wn2).text_frame.paragraphs)
    print("OK: icebreaker at [2], What's New labs bullet + notes updated")


if __name__ == "__main__":
    main()
