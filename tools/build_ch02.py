"""
build_ch02.py — refresh 1851 Ch02 'Quality Characteristics and Ethics in AI'
for revision A.2 (extend to GenAI).

Idempotent: always restores the carrier decks/1851-Ch02.pptx from
decks/_bak0/1851-Ch02.pptx before applying changes, so re-runs are safe.

Changes vs A.1 (29 slides -> 30 slides):
  CUT (redundant governance/culture duplicates):
    [18] Compliance Obligations for Developers  (dup of Documentation 22,
         Risk Mgmt 19, Governance 25)
    [24] Ethics Review in Project Lifecycle     (dup of impact assessments 19,
         approval gates 25)
    [26] Building Ethical Culture in Teams      (generic culture; substance
         absorbed by new 'Responsible GenAI Use in Software Teams')
  NEW (layout 'Content with Header. Full Page', speaker notes each):
    A Quality Characteristics of LLM-Based Features  (after quality block)
    B GenAI Ethics: New Failure Modes                (after bias/ethics block)
    C From Model Cards to System Cards               (after Model Cards)
    D Responsible GenAI Use in Software Teams        (governance section)
  UPDATE Summary slide: GenAI trust/ethics bullets alongside classic ones.

Run:  /Users/iahmad/Creator/Courses_and_conferences/LT/.venv-courseware/bin/python tools/build_ch02.py
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import slidekit as sk
import pptx_tools as pt

ROOT = Path(__file__).resolve().parent.parent
CARRIER = ROOT / "decks" / "1851-Ch02.pptx"
BAK = ROOT / "decks" / "_bak0" / "1851-Ch02.pptx"

# Sanity check: expected titles on the pristine carrier (0-based indices).
EXPECTED = {
    9: "Explainability in Practice",
    15: "Unintended Consequences of AI Optimization",
    18: "Compliance Obligations for Developers",
    23: "Model Cards and Data Sheets",
    24: "Ethics Review in Project Lifecycle",
    25: "Embedding Governance Into AI Workflows",
    26: "Building Ethical Culture in Teams",
    27: "Summary",
}

NEW_SLIDES = [
    # (title, bullets, speaker notes)
    (
        "Quality Characteristics of LLM-Based Features",
        [
            "Groundedness and faithfulness: output must be supported by the provided sources or verified data",
            "Hallucination rate: fabricated facts, references, and APIs—measure it with evals, never assume it",
            "Output consistency: equivalent answers across runs, seeds, and model versions",
            "Latency and cost are first-class quality attributes—they shape feasibility and user trust",
            "Classic attributes still apply, but flexibility, autonomy, and explainability are harder to guarantee",
        ],
        "Classic quality attributes like flexibility and explainability still matter, but "
        "LLM-based features add new first-class ones: groundedness, hallucination rate, and "
        "output consistency. These must be measured with evaluations, not assumed. Latency "
        "and per-call cost also become quality concerns because they shape which designs are "
        "feasible and whether users trust the feature.",
    ),
    (
        "GenAI Ethics: New Failure Modes",
        [
            "Hallucinated facts and code: confident, plausible, and wrong—invented APIs, citations, and configurations",
            "Licensing of generated code: models may reproduce licensed source; provenance and obligations are unclear",
            "Data leakage through prompts: source code, secrets, and personal data sent to external model providers",
            "Over-reliance and de-skilling: automation bias erodes review discipline and junior developer growth",
            "These failure modes compound classic risks (bias, opacity, reward hacking)—they do not replace them",
        ],
        "Generative systems fail in ways classic ML systems do not: they fabricate plausible "
        "facts and code, can emit licensed code with unclear provenance, and leak whatever "
        "data staff paste into prompts. Over-reliance is an ethical risk too—it de-skills "
        "teams and amplifies automation bias. Stress that these sit on top of bias and "
        "opacity, not instead of them.",
    ),
    (
        "From Model Cards to System Cards",
        [
            "Model cards document one model: training data, performance, and limitations",
            "A GenAI feature is a system: model + prompts + tools + data sources + guardrails",
            "System cards capture prompt templates, retrieval sources, tool permissions, and safety filters",
            "Include evaluation results: groundedness scores, hallucination rate, red-team findings, known failure modes",
            "Update on every prompt, model, or tool change—not only on retraining",
        ],
        "A model card describes a single model, but a GenAI feature is a whole system: model, "
        "prompts, tools, retrieval sources, and guardrails. A system card documents each of "
        "those plus evaluation results such as groundedness scores, hallucination rates, and "
        "red-team findings. It must be re-issued whenever prompts, models, or tools change, "
        "not just when a model is retrained.",
    ),
    (
        "Responsible GenAI Use in Software Teams",
        [
            "Usage policies: approved tools and models, permitted data classifications, no secrets in prompts",
            "Human review gates: generated code is reviewed like any external contribution—tests, security scan, license check",
            "Disclosure norms: mark AI-generated artifacts and record the tool and model versions used",
            "Accountability stays with the engineer: if you ship it, you own it",
            "Feed incidents and near-misses back into policy, training, and team norms",
        ],
        "Governance has to reach daily developer behavior. Teams need explicit usage "
        "policies—which tools are approved and what data may enter prompts—plus human review "
        "gates that treat generated code like any third-party contribution. Disclosure norms "
        "keep AI assistance visible, and accountability stays with the engineer who ships "
        "the work.",
    ),
]

SUMMARY_BULLETS = [
    "Effective AI systems require flexibility, explainability, and autonomy to support sustainable public service use",
    "LLM-based features add first-class quality attributes: groundedness, output consistency, latency, and cost",
    "Transparency builds institutional trust and supports legal rights (e.g., GDPR, DADM)",
    "Ethical risks span classic issues—bias, opaque logic, reward hacking—and GenAI failure modes like hallucinated output and prompt data leakage",
    "Developers must adopt mitigation tools like fairness dashboards, SHAP, and model auditing, plus human review gates for generated code",
    "Documentation evolves from model cards to system cards that capture prompts, tools, data sources, and evaluation results",
    "Governance workflows and team norms—usage policies, review gates, disclosure—make responsible AI a daily practice",
]

# Final order as 0-based indices into the slide list AFTER appending the four
# new slides (A=29, B=30, C=31, D=32) to the pristine 29-slide carrier.
# Slides 18, 24, 26 are simply not referenced, which drops them.
A, B, C, D = 29, 30, 31, 32
FINAL_ORDER = (
    [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    + [A]                       # after the quality-attributes block
    + [10, 11, 12, 13, 14, 15]
    + [B]                       # after the bias/ethics block
    + [16, 17, 19, 20, 21, 22, 23]
    + [C]                       # right after Model Cards and Data Sheets
    + [25]
    + [D]                       # governance section, after Embedding Governance
    + [27, 28]                  # Summary, then objectives recap
)


def main() -> None:
    # 1. Restore pristine carrier (idempotency).
    shutil.copy(BAK, CARRIER)

    prs = sk.open_prs(str(CARRIER))
    assert len(prs.slides) == 29, f"expected 29 pristine slides, got {len(prs.slides)}"
    for idx, title in EXPECTED.items():
        got = sk.slide_title(prs.slides[idx])
        assert got == title, f"slide {idx}: expected {title!r}, got {got!r}"

    layout = sk.pick_layout(prs, "Content with Header. Full Page")
    assert layout is not None, "layout 'Content with Header. Full Page' not found"

    # 2. Append the four new slides (append-before-arrange avoids partname
    #    collisions; slidekit also patches partname allocation).
    for title, bullets, notes in NEW_SLIDES:
        sk.add_bullets_slide(prs, title, bullets, notes=notes, layout=layout)

    # 3. Update the Summary slide (pristine index 27).
    sk.set_bullets(prs.slides[27], SUMMARY_BULLETS)

    # 4. Arrange: drop cuts, place new slides in their logical positions.
    pt.arrange(prs, FINAL_ORDER)

    sk.save(prs, str(CARRIER))

    # 5. Validate: reopen, print count + titles, assert no empty titles,
    #    no duplicate zip partnames, zip integrity.
    prs2 = sk.open_prs(str(CARRIER))
    n = len(prs2.slides)
    print(f"Saved {CARRIER.name}: {n} slides")
    titles = []
    for i, s in enumerate(prs2.slides):
        t = sk.slide_title(s)
        titles.append(t)
        print(f"{i + 1:>3} | {t}")
    assert n == 30, f"expected 30 slides, got {n}"
    assert all(t.strip() for t in titles), "empty title found"

    with zipfile.ZipFile(CARRIER) as z:
        names = z.namelist()
        dupes = {x for x in names if names.count(x) > 1}
        assert not dupes, f"duplicate zip partnames: {dupes}"
        assert z.testzip() is None, "zip integrity check failed"
    subprocess.run(["unzip", "-t", str(CARRIER)], check=True,
                   stdout=subprocess.DEVNULL)
    print("OK: 30 slides, no empty titles, no duplicate partnames, unzip -t passed")


if __name__ == "__main__":
    main()
