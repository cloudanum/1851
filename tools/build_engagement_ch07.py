"""build_engagement_ch07.py — engagement pass on decks/1851-Ch07.pptx (1851 A.2).

Idempotent: restores the deck from decks/_bak1/1851-Ch07.pptx (the post-build
A.2 snapshot), then inserts two NEW Do Now warm-ups:

  * "Do Now: Agent, Workflow, or Chatbot?" (Discussion layout) at index 7 —
    right after the early what-is-an-agent definition block
    (From Prompts to Agents / What Is an Agent? / The Cognitive Loop).
  * "Do Now: Design a Tool Schema" (Do Now Writing Offline layout) at index 14 —
    inside the tools/function-calling block, after "Tool Use / Function Calling
    Mechanics" and before the MCP slides.

Deck: 51 -> 53 slides.

Run (from 1851a2-author-input/):
    /Users/iahmad/Creator/Courses_and_conferences/LT/.venv-courseware/bin/python tools/build_engagement_ch07.py
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

SRC = ROOT / "decks" / "_bak1" / "1851-Ch07.pptx"
OUT = ROOT / "decks" / "1851-Ch07.pptx"

# --- NEW Do Now #1 (Discussion layout), inserted at index 7 ----------------------
DO_NOW_CLASSIFY = (
    "Do Now: Agent, Workflow, or Chatbot?",
    ["Time box: 6 minutes — classify each system: AGENT, WORKFLOW, or CHATBOT",
     "1. A Siri-style assistant that answers questions and sets timers on request",
     "2. A CI linter that flags style violations on every pull request",
     "3. An AutoGPT-style researcher that plans, searches the web, and writes a report",
     "4. A rules engine that routes insurance claims through if/then logic",
     "5. A Copilot-style assistant that suggests code completions in your IDE"],
    "Give individuals 6 minutes to classify all five, then compare answers popcorn-style. "
    "Key: 1 chatbot (answers on request, no goal loop); 2 workflow (fixed pipeline, no model "
    "deciding next steps); 3 agent (goal + tools + loop); 4 workflow (deterministic rules, no "
    "model at all); 5 chatbot as shipped — it only becomes agentic when it can run tools in a "
    "loop. Land the dividing line: who decides what happens next — a model in a loop, a fixed "
    "process, or a human between calls.",
)

# --- NEW Do Now #2 (Do Now Writing Offline layout), inserted at index 14 ----------
DO_NOW_SCHEMA = (
    "Do Now: Design a Tool Schema",
    ["Time box: 7 minutes — pen and paper, laptops closed",
     "Draft the JSON schema for a tool: query_issue_tracker(project, since, severity)",
     "Give the tool a name, a one-line description, and a parameters object",
     "Type every parameter; add an enum or format where it fits; mark what is required",
     "Swap with a neighbor and find one weakness in their schema"],
    "Students draft on paper, then swap and critique. A strong answer: name "
    "'query_issue_tracker'; description 'Search issues in a project opened since a date, "
    "filtered by severity'; parameters object with project (string, required), since (string, "
    "format: date), severity (string, enum: low/medium/high/critical). Debrief: a schema is a "
    "guardrail — types, enums, and required fields constrain the model before any code runs, "
    "which is exactly the discipline Lab 7.1 enforces with allow-lists.",
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
    assert len(prs.slides) == 51, f"expected 51 slides in _bak1 deck, has {len(prs.slides)}"

    # Insert higher position first so the lower index stays valid.
    # Tool-schema Do Now after "Tool Use / Function Calling Mechanics" (idx 13).
    add_engagement_slide(prs, "Do Now Writing Offline",
                         DO_NOW_SCHEMA[0], DO_NOW_SCHEMA[1], DO_NOW_SCHEMA[2],
                         at_index=14)
    # Classify Do Now after the definition block ("The Cognitive Loop", idx 6).
    add_engagement_slide(prs, "Discussion",
                         DO_NOW_CLASSIFY[0], DO_NOW_CLASSIFY[1], DO_NOW_CLASSIFY[2],
                         at_index=7)

    sk.save(prs, str(OUT))
    titles = validate(OUT, expect_slides=53)
    assert titles[7] == DO_NOW_CLASSIFY[0], f"unexpected slide at [7]: {titles[7]!r}"
    assert titles[15] == DO_NOW_SCHEMA[0], f"unexpected slide at [15]: {titles[15]!r}"
    print("OK: Do Now classify at [7], Do Now schema at [15]")


if __name__ == "__main__":
    main()
