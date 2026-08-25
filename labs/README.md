# Course 1851 (rev A.2) — VM Lab & Activity Pack

*AI for Software Engineers: Concepts and Techniques.* Audience: software
engineers. Each lab is a self-contained Jupyter notebook; the exercise slide
in the chapter deck points here ("Follow the detailed instructions in the Lab
X.Y notebook on your VM"). Activities 1.1 and 8.1 are print-and-run instruction
sheets (no code); Activities 4.2 and 8.2 are notebooks.

## Labs

| Lab | Title | Chapter | Day | Duration |
|-----|-------|---------|-----|----------|
| 1.1 | Assistant Recon: Where AI Helps, Where It Fails | Ch01 Introduction to AI in Software Development | 1 | ~30 min |
| 2.1 | Hallucination and Bias Audit | Ch02 Quality Characteristics and Ethics in AI | 1 | ~30 min |
| 3.1 | Decision Trees vs. an LLM | Ch03 Machine Learning Essentials | 1 | ~45 min |
| 4.1 | Building a PTCF Agent with the OpenAI API | Ch04 Prompt Engineering with OpenAI | 2 | ~45 min |
| 5.1 | TDD with an AI Pair | Ch05 Generative AI Across the SDLC | 2 | ~45 min |
| 6.1 | Write an Eval Suite | Ch06 Evaluating Generative AI Systems | 2 | ~45 min |
| 7.1 | Build a Guarded Agent | Ch07 AI Agents and Agentic Workflows | 3 | ~50 min |
| 8.1 | Red Team the Agent | Ch08 AI Security and Vulnerability Testing | 3 | ~40 min |
| 9.1 | Capstone: One Ticket, End to End | Ch09 Course Summary | 3 | ~50 min |

## Activities

| Activity | Title | Chapter | Format | Duration |
|----------|-------|---------|--------|----------|
| 1.1 | Turing Test Role-Play | Ch01 Introduction to AI in Software Development | instruction sheet (`activity_1.1_turing_test_roleplay.md`), trios | ~20 min |
| 4.2 | Prompt Pattern Clinic | Ch04 Prompt Engineering with OpenAI | notebook | ~30 min |
| 8.1 | Data Minimization Checklist | Ch08 AI Security and Vulnerability Testing | instruction sheet + `support_tickets_raw.csv`, pairs | ~30 min |
| 8.2 | Detect and Mask PII Using NLP | Ch08 AI Security and Vulnerability Testing | notebook (uses `support_tickets_raw.csv`) | ~40 min |

Do Now warm-ups are slide-contained (pen and paper, pairs, votes) and need no
VM assets. The one exception is the Chapter 8 ATLAS Matrix Do Now, which needs
a browser and internet access to `https://atlas.mitre.org/matrices/ATLAS`.

## Tutorial

| Tutorial | Title | When | Duration |
|----------|-------|------|----------|
| 0.1 | Getting Started: Your AI Workbench | Chapter 0, before Lab 1.1 | ~20 min |

`tutorial_0.1_getting_started.ipynb` is a guided first run of the shared
machinery: verifying the environment, live vs. mock mode, the five `course_ai`
calls, lab self-checks, and reset/key hygiene. Every lab and activity opens
with its own learning objectives; the tutorial covers what they all share.

## Setup

On the classroom VM everything below is **pre-provisioned** — open the
notebook and run. On your own machine:

```bash
pip install -r requirements.txt
export OPENAI_API_KEY=sk-...        # never paste the key into a notebook
python verify_setup.py              # 60-second smoke test
```

Launch Jupyter from the pack root (`jupyter lab` / `jupyter notebook`) so every
notebook sits next to `course_ai.py` and `support_tickets_raw.csv`.

The key is read from the environment (or a `.env` file at the course root)
and is never printed. The model is pinned via `OPENAI_MODEL`
(default `gpt-4o-mini`).

### Offline / no key: mock mode

Every notebook runs end-to-end with no key and no network:

```bash
export COURSE_AI_MOCK=1             # force the deterministic mock
```

Mock mode also engages automatically when `OPENAI_API_KEY` is absent. Mock
outputs are deterministic, clearly prefixed `[MOCK]`, and are behaviour-aware
for the exercises: Lab 5.1's mock generates a known-good (then refactorable)
`parse_semver`, Lab 6.1's mock produces both the baseline-quality and the
regressed "broken prompt" output so the suite demonstrably catches the
regression, and Lab 7.1's mock is a scripted planner that drives the full
agent loop (the HITL gate is fail-closed and auto-denies in non-interactive
runs). Labs whose exercise is to *score or audit the reply itself* pull fixed
transcripts from `course_ai.CANNED_*`: Lab 1.1's six probe replies include two
deliberate fluent partial failures, Lab 2.1's audit replies are fabricated or
skewed on purpose, Lab 3.1's LLM predictions are a fixed 8/10 list (just worse
than the decision tree), and Lab 9.1 ships a canned asset per capstone stage
(criteria, tests, implementation, blurb). Lab 8.1 contrasts a scripted
*unguarded* planner (Act 1 complies with the injection) with `course_ai`'s
scripted *guarded* planner (Act 3 refuses it). Where the mock cannot show the
phenomenon (e.g. temperature variance in Lab 4.1), a labelled canned transcript
from the instructor run stands in. Activity 8.2's mock is `CANNED_82`: the
per-ticket PII entity list a competent model returns — including the
contextual PII (names, a health note, a street address) that the lab's regex
pass demonstrably cannot see.

## Files

- `course_ai.py` — the one module between the learner and the SDK:
  `chat` / `chat_json` / `chat_tools` / `embed` / `cosine`, plus the
  deterministic mock and the `CANNED_*` transcripts. No keys, no snapshots;
  read it — it is short.
- `lab_*.ipynb` — the nine labs above.
- `tutorial_0.1_getting_started.ipynb` — the Day-1 orientation notebook above.
- `activity_1.1_turing_test_roleplay.md` — facilitator script: roles, rounds,
  question bank, AI-player script card, scorecard, debrief.
- `activity_4.2_prompt_clinic.ipynb` — the Ch04 activity.
- `activity_8.1_data_minimization_checklist.md` — checklist, worksheet, and
  instructor key, working on `support_tickets_raw.csv`.
- `activity_8.2_pii_detection.ipynb` — the Ch08 PII activity.
- `support_tickets_raw.csv` — twelve synthetic support tickets, shared by
  Activities 8.1 and 8.2 (all names, numbers, and addresses are fictional).
- `diagrams/` — figures cropped from the chapter decks and referenced by the
  notebooks (SDLC map, decision tree, PTCF blueprint, two-layer prompt
  architecture, agentic TDD loop, cognitive loop, injection kill chain,
  Zero-Trust agent layers). Keep the folder next to the notebooks.
- `requirements.txt`, `verify_setup.py` — environment spec and smoke test.
- Generated at runtime (safe to delete): `tdd_workspace/` (Lab 5.1),
  `capstone_workspace/` (Lab 9.1), `agent_repo/` (Labs 7.1 & 8.1),
  `eval_scorecard.json` (Lab 6.1), `audit_report.json` (Lab 2.1),
  `__pycache__/`, `.pytest_cache/`.

## Classroom VM provisioning notes

- Python 3.11+; `pip install -r requirements.txt`; then `python
  verify_setup.py` must pass.
- Inject `OPENAI_API_KEY` into the learner environment for live mode. With no
  key the pack still works everywhere — mock mode engages automatically.
- Internet allowlist: `api.openai.com` (live mode only) and
  `https://atlas.mitre.org` (Chapter 8 ATLAS Do Now, browser).
- Provision the pack as a single flat folder; learners launch Jupyter from
  that folder.

## Authoring QA

All notebooks are validated with `nbformat` and executed top-to-bottom with
`COURSE_AI_MOCK=1` and no key (nbclient). Re-run the same check after any
edit:

```bash
COURSE_AI_MOCK=1 OPENAI_API_KEY= jupyter nbconvert --to notebook --execute --inplace lab_X.Y_*.ipynb
```

(or the equivalent nbclient call). Never commit a real API key.
