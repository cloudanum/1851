# 1851 A.2 — Build Status (2026-08-24)

Major revision of course 1851 "AI for Software Engineers: Concepts and Techniques" (A.1, May 2025 → A.2, August 2026).

**Source of truth:** `1851/FROM_ftp/1851a1/` (published A.1 files from `/PDev/1851`, downloaded 2026-08-24). All decks rebuilt from `_bak0` snapshots via idempotent `tools/build_chNN.py` scripts (run with `../../../.venv-courseware/bin/python`; helpers: `tools/slidekit.py` + `tools/pptx_tools.py`).

## Themes delivered

- GenAI-first with OpenAI: new Ch04 (API anatomy, Responses API, 2026 model landscape) + API-driven labs.
- Agents & agentic workflows: new Ch07 (cognitive loop, ReAct, tools/function calling, MCP, PEV, HITL, multi-agent).
- GenAI across the SDLC incl. TDD: new Ch05 (10-slide TDD core) + Lab 5.1.
- Evals: new Ch06 (golden sets, rubric + LLM-as-Judge, eval pyramid, Eval-Driven Development, evals in CI) + Lab 6.1.
- Traditional ML cut 178 → 58 slides (**67% reduction**), condensed into Ch03 "Machine Learning Essentials" with 4 new transformer/LLM bridge slides.
- Diagrams/slides reused from courses 2016 (agent anatomy, PEV, determinism sandwich, zero-trust agent), 1258 (model landscape, prompt figures, Ragas).

## Decks (all validated: reopen, titles, no dup partnames, `unzip -t`)

| Deck | Slides | Was | Change |
|---|---|---|---|
| Ch00 Introduction and Overview | 4 | 3 | new objectives/TOC, "What's New" slide |
| Ch01 Introduction to AI in Software Development | 38 | 48 | 2026 rewrite; Activity 1.1 kept; new Activity 1.2 |
| Ch02 Quality Characteristics and Ethics in AI | 30 | 29 | +4 GenAI trust/ethics slides |
| Ch03 Machine Learning Essentials | 58 | 54(+124) | absorbs old Ch04/05/07; Activity 3.1 kept |
| Ch04 Prompt Engineering with OpenAI | 48 | (old Ch06, 37) | +OpenAI API block, 2026 prompting; Lab 4.1/4.2 |
| Ch05 Generative AI Across the SDLC | 50 | (old Ch04 shell) | NEW — authored; Lab 5.1 |
| Ch06 Evaluating Generative AI Systems | 44 | (old Ch05 shell) | NEW — authored + EDD/TEVV clones; Lab 6.1 |
| Ch07 AI Agents and Agentic Workflows | 51 | (old Ch07 shell) | NEW — 42 clones + 4 authored; Lab 7.1 |
| Ch08 AI Security and Vulnerability Testing | 52 | 51 | +agentic threats; Activities 8.1/8.2 kept |
| Ch09 Course Summary | 3 | 2 | new takeaways + objectives recap |
| **Total** | **378** | 348 | |

## Labs (`labs/`)

Nine labs + one activity: `lab_1.1_assistant_recon.ipynb`, `lab_2.1_hallucination_bias_audit.ipynb`, `lab_3.1_decision_tree_vs_llm.ipynb`, `lab_4.1_openai_api.ipynb`, `activity_4.2_prompt_clinic.ipynb` (renamed from `lab_4.2_prompt_studio.ipynb`; title now "Activity 4.2: Prompt Pattern Clinic"), `lab_5.1_tdd_ai_pair.ipynb`, `lab_6.1_eval_suite.ipynb`, `lab_7.1_guarded_agent.ipynb`, `lab_8.1_red_team_the_agent.ipynb`, `lab_9.1_capstone_ticket.ipynb`; `course_ai.py` (OpenAI wrapper, `COURSE_AI_MOCK=1` offline mode, plus `CANNED_*` deterministic transcripts for the scoring/audit labs 1.1/2.1/3.1/9.1), `README.md`. All ten notebooks validated (nbformat + ast.parse) and executed top-to-bottom in mock mode with zero errors (nbclient). Live-key pass still recommended before release (Responses API cells target openai SDK 2.44).

## Instructor Guide

`docs/1851-IGa2.docx` via `tools/build_ig.py`: new A.2 Course History paragraph + fully populated 3-day timeline (was placeholders in A.1).

## Known follow-ups (non-blocking)

- New authored slides are static text on LT layouts — a light visual polish pass in PowerPoint is worthwhile (esp. cloned diagram slides Ch07 s7/8/15/26/35, Ch08 s36-41).
- A few cloned slides keep source-course speaker-note references (2016/1258 chapter numbers) — instructor-facing only.
- Ch07 s27 case study is SecOps-flavored (SOC) — content accurate, flagged for author proof.
- `!1851-FPa1.pptx` combined deck intentionally not rebuilt — chapter decks are authoritative; Pubs recombines.
- Activity 3.1 and 8.1/8.2 still reference VM instructions (unchanged from A.1; VM provisioning is LT-side).

## Layout fix pass (2026-08-24)

- Geometry: repaired 144 text placeholders to exact layout geometry (Ch00 1, Ch01 20, Ch02 7, Ch03 18, Ch04 7, Ch05 50, Ch06 7, Ch07 2, Ch08 32) — 57 zero-width `cx="0"` bugs (invisible text: Ch04 4, Ch05 48, Ch06 5) + 87 stretched/inheriting bodies normalized. 6 more bodies were deliberately left at custom geometry because the layout geometry would overlap a picture (Ch01 s31, Ch02 s15, Ch03 s19, Ch04 s41, Ch08 s10, Ch08 s11).
- Text fit: trimmed 5 Do Now bodies to ≤5 lines (Ch04 s12, Ch04 s33, Ch05 s27, Ch07 s8, Ch08 s40); speaker notes untouched.
- Empty slides: filled Ch03 s20 'K-Means Clustering Algorithms' (was title-only). No husk deletions were needed: Ch01 s26 'The GenAI-Augmented SDLC' and Ch06 s22 'Choosing the Right Evaluation Method' already carry their tables — the title-only duplicates from the diagnosis do not exist in the current decks. Ch03 s18 'Supervised vs. Unsupervised ML' already had its comparison table and was left untouched. Deck slide counts unchanged (course total 400).
- Note: `tools/fix_layout_pass.py` is the mandatory final step after any deck rebuild (earlier build scripts contain the height-resize idiom that causes cx=0). It is idempotent (verified: runs 2 and 3 made 0 changes, acceptance PASS).

## Ch03 invisible-body fix + diagram enrichment (2026-08-24, second pass)

- Root cause: 5 slides (s38 'Why Evaluation Matters', s45 'Business-Driven Metric Selection', s50 'Forward Propagation', s51 'Backpropagation', s52 'Neural Network Training Loop') had body placeholders with NO xfrm extent at all (`<a:ext>` missing, not `cx="0"`), so the earlier `fix_layout_pass.py` missed them — they rendered as title + stock photo with invisible bullets.
- Fix: `tools/fix_ch03_empty.py` (idempotent) restores body geometry and inserts one SVG-authored diagram per slide (sources + 1500px PNG renders in `tools/diagrams/`): evaluation feedback loop, FN/FP cost → metric selection matrix, forward-propagation network, backpropagation gradient flow, 5-step training loop.
- Verified programmatically: 0 zero-width text shapes remain in Ch03; bullets/diagram bands do not overlap (bullets bottom 3.14in, diagrams 3.30–6.72in, footer at 6.89in).
- Ch03 re-uploaded to `/PDev/1851/1851a2/` (10,417,327 bytes, matches local md5).

## Missing-extent repair + Ch04 agent-prompting enrichment (2026-08-24, third pass)

- Root cause: a second variant of the build bug — body placeholders with NO `<a:ext>` element (python-pptx `width is None`), invisible to the first fix pass which matched `cx="0"`. Corrected acceptance scan (text-bearing shapes with `width is None or < 10000`, deck-wide) found 65 slides: Ch04 17, Ch06 5, Ch07 41, Ch08 2. All repaired by `tools/fix_missing_extent.py` (idempotent, built-in acceptance; geometry rule: right-picture → left column, else full width). Post-repair overlap check: only 3 pre-existing overlaps, none introduced.
- Ch04 enriched from `30_Agents_Ch_03_The Art of Agent Prompting.pdf` via `tools/build_ch04_pdf_enrichment.py` (idempotent): 51→56 slides — new "The Two-Layer Prompt Architecture" (s22), "Designing Thinking Agents: The Prompting Spectrum" (s29), "The PTCF Blueprint" (s37), "PTCF Walk-Through: An Enterprise Billing Agent" (s38), "From Instructions to Constitutions" (s49); Agenda/bridge/Summary lines added; Do Now s12 → "Constitution or Command?", Do Now s35 → PTCF misalignment clinic (answer keys in speaker notes).
- New diagrams (SVG→PNG in `tools/diagrams/`, visually verified): two-layer-architecture, ptcf-blueprint, thinking-spectrum.
- Labs rebuilt from the PDF (`tools/build_lab41_activity42.py`): Lab 4.1 → "Building a PTCF Agent with the OpenAI API" (assemble constitution, stimuli, ablation, temperature sweep); Activity 4.2 → PTCF anti-pattern clinic (3 bad→fixed cases + peer rubric). Both execute clean in mock mode (nbclient). Labs README + IG timeline retitled; `1851-Labs-a2.zip` and `1851-IGa2.docx` rebuilt.
- Final acceptance: all 10 decks 0 invisible-text shapes; Ch04 56 slides; course total 405.
