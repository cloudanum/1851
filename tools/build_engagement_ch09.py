"""build_engagement_ch09.py — engagement pass on decks/1851-Ch09.pptx (1851 A.2).

Idempotent: restores the deck from decks/_bak1/1851-Ch09.pptx (the post-build
A.2 snapshot), then
  * inserts a NEW "Lab 9.1: Capstone — One Ticket, End to End" (Exercise
    Reference Slide with Typing Hands layout) at index 2 — right after the
    Course Summary takeaways slide,
  * appends a NEW "Do Now: Your Monday Experiment" (Do Now Writing Offline
    layout) at the very end, after the Course Objectives recap.

Deck: 3 -> 5 slides.

Run (from 1851a2-author-input/):
    /Users/iahmad/Creator/Courses_and_conferences/LT/.venv-courseware/bin/python tools/build_engagement_ch09.py
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

SRC = ROOT / "decks" / "_bak1" / "1851-Ch09.pptx"
OUT = ROOT / "decks" / "1851-Ch09.pptx"

# --- NEW Lab 9.1 (Exercise Reference Slide with Typing Hands), at index 2 --------
LAB_9_1 = (
    "Lab 9.1: Capstone — One Ticket, End to End",
    ["Follow the detailed instructions in the Lab 9.1 notebook on your VM",
     "Take one mini-ticket end to end: acceptance criteria → tests first → "
     "AI implementation → eval gate → docs",
     "You write the criteria and the tests; the AI drafts the implementation",
     "Scorecard your own process: where the AI saved time, and where it cost you",
     "Time box: about 50 minutes"],
    "The capstone: one small ticket carried through the entire GenAI workflow from the "
    "course, with tests and the eval gate as the quality bars. Keep it to about 50 minutes "
    "and push students to pick a genuinely small ticket — the point is the full loop, not "
    "the feature. Debrief with the scorecard: two or three volunteers share where the AI "
    "helped most and where it slowed them down.",
)

# --- NEW Do Now (Do Now Writing Offline layout), appended at the very end --------
DO_NOW_MONDAY = (
    "Do Now: Your Monday Experiment",
    ["Time box: 5 minutes — pen and paper",
     "Write down ONE GenAI technique from this course you will try on real work next week",
     "Name the concrete task you will try it on",
     "Decide how you will measure it: time saved, defects caught, review comments, eval score",
     "Finish the sentence: \u2018I will know it worked if ...\u2019"],
    "Close the course with a written commitment: one technique, one real task, one concrete "
    "measure. Give the full 5 minutes — the measure ('I will know it worked if ...') is the "
    "hard part. Collect the cards or have a few volunteers read theirs out; a public "
    "commitment is far more likely to happen. Point students at the course resources for "
    "when they get stuck.",
)


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


# --- validation ------------------------------------------------------------------
def validate(path: Path, expect_slides: int) -> int:
    prs = sk.open_prs(path)
    n = len(prs.slides)
    assert n == expect_slides, f"{n} slides != expected {expect_slides}"
    print(f"\n{path}: {n} slides")
    titles = []
    for i, s in enumerate(prs.slides):
        title = sk.slide_title(s)
        titles.append(title)
        print(f"  [{i}] {title!r}")
        assert title.strip(), f"slide {i} has an empty title"
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    dupes = sorted({nm for nm in names if names.count(nm) > 1})
    assert not dupes, f"duplicate zip partnames: {dupes}"
    r = subprocess.run(["unzip", "-t", str(path)], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    print("  zip partnames unique:", len(names), "parts;", r.stdout.strip().splitlines()[-1])
    return titles


def main() -> None:
    shutil.copyfile(SRC, OUT)  # idempotent: always start from the _bak1 snapshot
    prs = sk.open_prs(OUT)
    assert len(prs.slides) == 3, f"expected 3 slides in _bak1 deck, has {len(prs.slides)}"

    # Lab 9.1 after the takeaways slide (idx 1) -> lands at idx 2; the closing
    # Do Now appends at the very end (after the objectives recap).
    add_engagement_slide(prs, "Exercise Reference Slide with Typing Hands",
                         LAB_9_1[0], LAB_9_1[1], LAB_9_1[2], at_index=2)
    add_engagement_slide(prs, "Do Now Writing Offline",
                         DO_NOW_MONDAY[0], DO_NOW_MONDAY[1], DO_NOW_MONDAY[2])

    sk.save(prs, str(OUT))
    titles = validate(OUT, expect_slides=5)
    assert titles[2] == LAB_9_1[0], f"unexpected slide at [2]: {titles[2]!r}"
    assert titles[4] == DO_NOW_MONDAY[0], f"unexpected slide at [4]: {titles[4]!r}"
    print("OK: Lab 9.1 at [2], Monday-experiment Do Now at [4]")


if __name__ == "__main__":
    main()
