# Course 1851 — AI for Software Engineers (rev A.2)

Code and tooling for the Learning Tree course *1851 — AI for Software
Engineers: Concepts and Techniques*, revision A.2 (August 2026).

## Layout

- `labs/` — the validated VM lab & activity pack: Tutorial 0.1, nine labs,
  four activities, `course_ai.py` (the SDK wrapper with a deterministic
  offline mock), the shared support-tickets dataset, deck diagrams, and
  `verify_setup.py`. Ships to the classroom VM as `1851-VM-assets-a2.zip`;
  see `labs/README.md` for setup, mock mode, and provisioning notes.
- `tools/` — the python-pptx build/fix scripts that generate the A.2 chapter
  decks and instructor guide (`build_chNN.py`, `build_engagement_chNN.py`,
  `slidekit.py`, `pptx_tools.py`, layout/schema repair passes,
  `build_ch05_agents.py`, `apply_notes.py`, `validate_pack.py`), plus their
  figure inputs (`tools/diagrams/`, `tools/ch05_assets/`).
- `MOVES.md` — A.1 → A.2 change manifest for Publications.
- `STATUS.md` — A.2 build-status snapshot from the authoring workspace.
- `a1-legacy/` — superseded A.1 (May 2025) labs, activities, and datasets,
  kept for reference.

## Conventions

- No slide decks, instructor-guide binaries, PDFs, or zip artifacts in git —
  those ship through Learning Tree's PDev/FTP channels.
- Every notebook runs end-to-end offline: `COURSE_AI_MOCK=1` (or simply no
  `OPENAI_API_KEY`) engages a deterministic mock whose outputs are prefixed
  `[MOCK]`. The model is pinned via `OPENAI_MODEL` (default `gpt-4o-mini`).
- QA: from `labs/`, run `python tools/validate_pack.py` from the repo root's
  sibling tooling (see `labs/README.md` → Authoring QA) — all notebooks are
  executed top-to-bottom in mock mode before any release.
