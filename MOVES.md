# 1851 A.2 — MOVES (change manifest for Publications)

Old → new mapping. All A.1 files were taken from `/PDev/1851/1851a1` (downloaded 2026-08-24).

## Chapter decks

| A.1 file (title) | A.2 file (title) | Operation |
|---|---|---|
| 1851-Ch00.pptx (Introduction and Overview) | 1851-Ch00.pptx (same) | REVISED — new course objectives + contents, added "What's New in This Revision" |
| 1851-Ch01.pptx (Introduction to AI in Software Development) | 1851-Ch01.pptx (same) | REVISED — 2026 GenAI-first rewrite (48→38); Activity 1.1 kept, Activity 1.2 added |
| 1851-Ch02.pptx (Quality Characteristics and Ethics in AI) | 1851-Ch02.pptx (same) | REVISED — extended to GenAI (29→30) |
| 1851-Ch03.pptx (Machine Learning Overview) | 1851-Ch03.pptx (Machine Learning Essentials) | REBUILT — condensed to essentials; absorbs kept content from A.1 Ch04/Ch05/Ch07 (178→58 slides total ML) |
| 1851-Ch06.pptx (Prompt Engineering for Generative AI) | 1851-Ch04.pptx (Prompt Engineering with OpenAI) | MOVED+EXPANDED (37→48) — OpenAI API anatomy, 2026 model landscape, Labs 4.1/4.2 |
| — (new) | 1851-Ch05.pptx (Generative AI Across the Software Development Lifecycle) | NEW (50) — incl. TDD-with-AI core; Lab 5.1. (A.1 Ch04 "Data Preparation and Handling" RETIRED; kept content lives in Ch03) |
| — (new) | 1851-Ch06.pptx (Evaluating Generative AI Systems) | NEW (44) — evals, LLM-as-Judge, Eval-Driven Development; Lab 6.1. (A.1 Ch05 "Model Evaluation Metrics" RETIRED; kept content lives in Ch03) |
| — (new) | 1851-Ch07.pptx (AI Agents and Agentic Workflows) | NEW (51) — anatomy, tools, MCP, patterns, workflows; Lab 7.1. (A.1 Ch07 "Neural Networks Introduction" RETIRED; kept content lives in Ch03) |
| 1851-Ch08.pptx (AI Security and Vulnerability Testing) | 1851-Ch08.pptx (same) | REVISED — agentic-era threats added (51→52); Activities 8.1/8.2 kept |
| 1851-Ch09.pptx (Course Summary) | 1851-Ch09.pptx (same) | REVISED — new takeaways + objectives recap (2→3) |
| !1851-FPa1.pptx (combined deck) | — | SUPERSEDED — chapter decks are authoritative |

## Other files

| A.1 | A.2 | Operation |
|---|---|---|
| 1851-IGa1.docx | 1851-IGa2.docx | REVISED — new Course History paragraph (A.2, August 2026); Timeline table fully populated (was placeholders) |
| MyLearningGoals.pdf | MyLearningGoals.pdf | UNCHANGED |
| — | 1851-Labs-a2.zip | NEW — 9 Jupyter labs (1.1, 2.1, 3.1, 4.1, 5.1, 6.1, 7.1, 8.1, 9.1) + Activity 4.2 (Prompt Pattern Clinic) + `course_ai.py` support module + README. Labs use `OPENAI_API_KEY` with a deterministic offline mock (`COURSE_AI_MOCK=1`) |

## Exercise numbering

- Kept: Activity 1.1 (Turing Test role-play), Activity 8.1 (Data Minimization Checklist), Activity 8.2 (Detect and Mask PII).
- Promoted: Activity 3.1 (Decision Tree Classifier, VM activity) → Lab 3.1 (Decision Trees vs. an LLM) — slide retitled, new notebook supersedes the old VM instructions.
- Removed: Activity 3.2 (NER), Activity 3.3 (K-Means), Ch03 "Do Now" informal exercises.
- Added: Activity 1.2 (Map Your Own SDLC); Labs 1.1 (Assistant Recon), 2.1 (Hallucination and Bias Audit), 3.1 (Decision Trees vs. an LLM), 4.1 (Building a PTCF Agent with the OpenAI API), 5.1 (TDD with an AI Pair), 6.1 (Write an Eval Suite), 7.1 (Build a Guarded Agent), 8.1 (Red Team the Agent), 9.1 (Capstone: One Ticket, End to End).
- Renamed: Lab 4.2 (Prompt Pattern Studio) → Activity 4.2 (Prompt Pattern Clinic), file `lab_4.2_prompt_studio.ipynb` → `activity_4.2_prompt_clinic.ipynb`.
- Note for slide cross-references: Lab 8.1 shares its number with kept Activity 8.1 (different exercise types; the deck's Activity numbering is unchanged).

## A.2 post-submission enrichment (Aug 24, 2026)

- Ch04 expanded 51→56 slides with agent-prompting material (from the companion book chapter "The Art of Agent Prompting"): new slides "The Two-Layer Prompt Architecture" (s22), "Designing Thinking Agents: The Prompting Spectrum" (s29), "The PTCF Blueprint" (s37), "PTCF Walk-Through: An Enterprise Billing Agent" (s38), "From Instructions to Constitutions" (s49); Do Now s12 retitled "Do Now: Constitution or Command?"; Do Now s35 rebuilt as a PTCF clinic; Lab 4.1 retitled "Building a PTCF Agent with the OpenAI API" (notebook rebuilt around PTCF); Activity 4.2 notebook rebuilt around PTCF anti-pattern cases.

## A.2 post-submission fixes and enrichment (Aug 24, 2026, PM)

- Ch05 expanded 52→64 slides with a "Software Development Agents" section (from the companion book chapter): three agent classes, tooling stack, adoption maturity, test-driven generation, generate–test–refine loop, multi-agent feature teams, compliance-driven agents + PCI DSS case study, self-improving agents + governance, new "Do Now: How Much Autonomy?" discussion. Agenda, objectives, summary, and s46 pointer updated; 5 diagrams included.
- All decks: repaired schema-invalid OOXML that triggered Windows PowerPoint's "repair" prompt (reversed headEnd/tailEnd in connector lines, negative shape extents, misplaced paragraph properties, illegal xml:space attributes, non-schema buSzPct values). All parts now validate against ECMA-376 transitional schemas.
- Ch07: fixed 11 mispositioned connector arrows on slides 7, 9, 26 (incl. the Planner-Executor-Verifier feedback edge that ran off-canvas).
- All decks: speaker notes added to every slide that lacked them (52 slides); all 466 slides now have notes.
- Verified: no slide references a workbook; all labs/activities point to the VM notebooks.

## A.2 slide-repair pass (Aug 24, 2026, late PM)

- Ch03 s20 (K-Means): removed legacy A.1 steps-table image that overlapped the new bullets.
- Ch03 s39/s40 (Regression Metrics): removed broken/non-rendering header photos, restored titles from invisible white to template blue, moved bullets up to standard position.
- Ch03 s34 (Random vs. Stratified Splitting): removed non-rendering orphan decorative image + empty placeholder; content widened to full page.
- Ch08 s26 (Data Minimization in Practice): removed 7 empty padding paragraphs that pushed the "Reflection" bullet behind the table image; table moved down for clean separation.
- Ch02 s15 (Bias in Data and Models): title changed from low-contrast white to template blue (0055B8) on the full-bleed photo.
- All decks: stray leading/trailing whitespace trimmed from 17 slide titles.
- Full visual QA sweep of all 466 slides (rendered + pixel-scanned): no remaining blank titles, blank bodies, or text/image collisions. All decks re-validated against ECMA-376 transitional schemas — clean.

## A.2 VM lab/activity pack (Aug 24, 2026, evening)

- `1851-VM-assets-a2.zip` (`1851-VM-pack/`): complete, validated VM package — 9 labs, 4 activities, `course_ai.py`, `requirements.txt`, `verify_setup.py` smoke test, rewritten README with setup, reset, and provisioning notes.
- New assets authored (slides referenced them "in the VM" but none existed): Activity 1.1 Turing Test Role-Play facilitator script (roles, rounds, question bank, AI-player script card, scorecard, debrief); Activity 8.1 Data Minimization Checklist + shared `support_tickets_raw.csv` (12 synthetic tickets, incl. instructor key); Activity 8.2 PII detection notebook (regex pass vs. LLM semantic pass via `chat_json`, masking with re-scan proof; mock served by new `course_ai.CANNED_82` per-ticket entity transcript).
- `requirements.txt` adds matplotlib (Lab 7.1's agent-loop diagram) — the old README's pip line missed it.
- Do Now audit (all 19, full slide text): all slide-contained; only Ch08's ATLAS Matrix Do Now needs VM internet (atlas.mitre.org) — documented as a provisioning requirement, no asset needed.
- QA: all 11 notebooks executed top-to-bottom with `COURSE_AI_MOCK=1` and no key — three consecutive clean passes (twice in place, once from the shipped zip); `verify_setup.py` passes.

## A.2 VM pack enrichment (Aug 24, 2026, late evening)

- New `tutorial_0.1_getting_started.ipynb` (Chapter 0, ~20 min, before Lab 1.1): guided first run of the shared machinery — environment check, live vs. mock mode, all five `course_ai` calls (`chat`, `chat_json`, `embed`+`cosine`, one `chat_tools` agent preview), lab self-checks, reset and key hygiene.
- Learning objectives added to the four activity assets (Activities 1.1, 8.1 instruction sheets; 4.2, 8.2 notebooks) in the labs' house format — the nine labs already had them.
- Fixed a latent live-mode bug in `course_ai.chat_tools`: the live path returned the message *content string* instead of the message object, so `msg.tool_calls` would have failed on a keyed run (mock mode masked it). README gained a Tutorial section and Files entry.
- QA: all 12 notebooks executed top-to-bottom in mock mode — two clean passes in place plus one from the re-shipped zip; `verify_setup.py` passes.

## A.2 VM pack diagrams (Aug 24, 2026, night)

- New `1851-VM-pack/diagrams/` — eight figures cropped at 200 DPI from the final QA renders of the chapter decks, embedded in the notebooks at the section they illustrate: cognitive loop (Ch07 s7 → Tutorial 0.1 Step 5), PTCF blueprint (Ch04 s37 → Lab 4.1 Step 1 and Activity 4.2), two-layer prompt architecture (Ch04 s22 → Lab 4.1 Step 2), agentic TDD loop (Ch05 s51 → Lab 5.1 Step 3), decision tree (Ch03 s13 → Lab 3.1 Step 3), Zero-Trust agent layers (Ch08 s42 → Lab 7.1 Stage C), injection kill chain (Ch08 s37 → Lab 8.1 Act 1), GenAI-augmented SDLC map (Ch01 s26 → Lab 9.1 intro). Every crop visually reviewed; no clipped content.
- QA: all 12 notebooks re-validated in mock mode (clean pass in place + one from the re-shipped zip); image-link check — no broken references in the shipped zip.
