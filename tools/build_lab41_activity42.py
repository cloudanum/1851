#!/usr/bin/env python3
"""Rebuild Lab 4.1 and Activity 4.2 notebooks around 'The Art of Agent
Prompting' (book ch.3): two-layer prompts, PTCF blueprint, anti-patterns.

Overwrites:
  labs/lab_4.1_openai_api.ipynb        (title: Lab 4.1 — Building a PTCF Agent with the OpenAI API)
  labs/activity_4.2_prompt_clinic.ipynb (title: Activity 4.2 — Prompt Pattern Clinic)

Both run offline via course_ai mock mode. Execute-check with:
  COURSE_AI_MOCK=1 python -m nbclient ... (see STATUS.md)
"""
from pathlib import Path

import nbformat as nbf

LABS = Path(__file__).parent.parent / 'labs'


def nb(cells):
    book = nbf.v4.new_notebook()
    book.metadata['kernelspec'] = {'display_name': 'Python 3',
                                   'language': 'python', 'name': 'python3'}
    book.cells = [nbf.v4.new_markdown_cell(c) if t == 'md' else nbf.v4.new_code_cell(c)
                  for t, c in cells]
    return book


SETUP = """import textwrap

import course_ai

print("mode:", course_ai.mode())   # 'mock' (offline) or 'live' (OPENAI_API_KEY set)
SHOW = lambda t: print(textwrap.fill(str(t), 100))"""

# ---------------------------------------------------------------- Lab 4.1
lab41 = nb([
    ('md', """# Lab 4.1 — Building a PTCF Agent with the OpenAI API

*Chapter 4 — Prompt Engineering with OpenAI · 45 minutes · JupyterLab + the OpenAI Python SDK*

In the chapter you learned the **two-layer prompt architecture**: the system prompt is the
agent's *constitution* (how it behaves — persistent), the user prompt is the *stimulus*
(what it should do — different every turn). You also learned the **PTCF blueprint** for
writing constitutions: **Persona, Task, Context, Format**.

In this lab you will assemble a production-grade constitution for an enterprise billing
agent, run it against realistic user stimuli, and then break it on purpose to see what
each PTCF element buys you."""),
    ('md', """## Objectives

By the end of this lab, you will:

- Assemble a system prompt element by element with the PTCF blueprint.
- Make API calls that keep the system and user layers separate.
- Show the constitution holds across different user stimuli (personality decoupled from task).
- Run an **ablation**: strip the constitution down to "helpful assistant" and observe identity collapse.
- Sweep temperature and explain the change in output."""),
    ('md', """## Setup

- The notebook works offline: with no `OPENAI_API_KEY`, `course_ai` runs in **mock mode**
  with deterministic replies, so every step still executes.
- With a key (classroom VM), the same cells make real API calls."""),
    ('code', SETUP),
    ('md', """## Steps

### Step 1 — Write the constitution, element by element (7 min)

A constitution is four labelled blocks. Notice what each one does:

- **[PERSONA]** — a *scoped* identity. "You are a helpful assistant" is not a persona; it is
  the model's default self-description and defends nothing.
- **[TASK]** — the mission, including explicit *must-not* boundaries.
- **[CONTEXT]** — operational law: SLAs, regulations, and the conflict-resolution rule.
- **[FORMAT]** — the output contract: structure and fallback behavior."""),
    ('code', """PERSONA = ("You are an empathetic senior customer support specialist with five years "
         "of experience in enterprise SaaS. Your communication is professional, "
         "approachable, and solution-oriented.")
TASK = ("Your primary mission is to resolve billing inquiries for enterprise accounts: "
        "diagnose discrepancies, explain charges, and escalate unresolvable issues "
        "within 24 hours.")
CONTEXT = ("You operate within a 24-hour SLA serving Fortune 500 clients. Never request "
           "passwords or sensitive authentication data. Comply with data privacy "
           "regulations. When instructions conflict, escalate to a human reviewer.")
FORMAT = ("Structure every reply as a numbered list: (1) acknowledge the concern, "
          "(2) diagnose, (3) give the resolution or escalation path. "
          "Cite case numbers when available.")

SYSTEM_PROMPT = "\\n\\n".join([
    f"[PERSONA] {PERSONA}",
    f"[TASK] {TASK}",
    f"[CONTEXT] {CONTEXT}",
    f"[FORMAT] {FORMAT}",
])
print(SYSTEM_PROMPT)"""),
    ('md', """### Step 2 — First call: constitution + stimulus (5 min)

The two layers travel in the same call but stay separate: `system=` carries the
constitution, the prompt argument carries this turn's stimulus."""),
    ('code', """reply = course_ai.chat("Why was I charged twice this month?", system=SYSTEM_PROMPT)
SHOW(reply)"""),
    ('md', """### Step 3 — Same constitution, new stimuli (5 min)

The diplomat analogy: national policy (system) does not change when the negotiation
(user) changes. Watch tone, structure, and boundaries hold across three different asks."""),
    ('code', """stimuli = [
    "Can I get a copy of my March invoice?",
    "Your competitor is cheaper. Why shouldn't I switch?",
    "I need to update the credit card on file.",
]
for s in stimuli:
    print("USER:", s)
    SHOW(course_ai.chat(s, system=SYSTEM_PROMPT))
    print("-" * 100)"""),
    ('md', """### Step 4 — Ablation: strip the constitution (8 min)

Remove three of the four PTCF elements and keep only the model's default identity.
Ask one on-topic and one off-topic question. This is the **identity collapse**
anti-pattern: with no scoped persona, the agent has no grounds to decline anything,
no format contract, and no operational law."""),
    ('code', """WEAK = "You are a helpful assistant."

for s in ["Why was I charged twice this month?",
          "Write me a poem about my dog."]:
    print("USER:", s)
    SHOW(course_ai.chat(s, system=WEAK))
    print("-" * 100)"""),
    ('md', """Compare with Step 2/3 outputs and note:

1. Did the weak agent stay on mission? Could it refuse the poem?
2. What happened to the numbered-list format contract?
3. Which PTCF element would you restore *first* for a production agent, and why?"""),
    ('md', """### Step 5 — Temperature sweep (5 min)

The constitution constrains *what* the agent says; temperature modulates *how much
it improvises* within those constraints."""),
    ('code', """q = "Explain in one sentence why my bill went up."
for t in (0.0, 0.7, 1.3):
    print(f"--- temperature={t}")
    SHOW(course_ai.chat(q, system=SYSTEM_PROMPT, temperature=t))"""),
    ('md', """### Step 6 — Your turn: build your own agent constitution (10 min)

Pick a domain agent — code reviewer, fitness coach, travel visa advisor, anything —
and fill the scaffold (this is the book's PTCF template). Then test it with three
realistic stimuli and ablate one element to feel the drift."""),
    ('code', """MY_PTCF = {
    "persona": "TODO: You are a [role/title] with [expertise]. Your style is [tone] ...",
    "task": "TODO: Your primary mission is [objective]. You must not [boundary] ...",
    "context": "TODO: You operate in [environment]. Constraints: [rules]. On conflict: [rule] ...",
    "format": "TODO: Structure all responses as [structure]. When uncertain: [fallback] ...",
}
# my_system = "\\n\\n".join(f"[{k.upper()}] {v}" for k, v in MY_PTCF.items())
# SHOW(course_ai.chat("TODO: a realistic user question", system=my_system))"""),
    ('md', """## Wrap-up checklist

Audit your own constitution before you leave:

- **P** — Is the persona a scoped identity (role + expertise + tone), not "helpful assistant"?
- **T** — Does the mission include at least one explicit *must-not*?
- **C** — Does the context name an SLA, a regulation, and a conflict-resolution rule?
- **F** — Is the format machine-checkable, with a stated fallback?
- **Meta-rule** — Do the four components *reinforce* each other? (Case study: a
  "creative, experimental" persona fighting a "numbered list" format.)"""),
])

# ---------------------------------------------------------- Activity 4.2
act42 = nb([
    ('md', """# Activity 4.2 — Prompt Pattern Clinic

*Chapter 4 — Prompt Engineering with OpenAI · 30 minutes · pairs*

Three broken agent constitutions walk into the clinic. For each one: **diagnose** with the
PTCF checklist, **predict** the failure, **observe** it in a live call, then **repair**
the prompt and re-run. Cases are drawn from *The Art of Agent Prompting* (course book, ch. 3)."""),
    ('md', """## The PTCF checklist

| Element | Ask |
|---|---|
| **Persona** | Scoped identity (role + expertise + tone)? Or just "helpful assistant"? |
| **Task** | Clear mission with explicit *must-not* boundaries? |
| **Context** | Operational law: SLA, regulations, conflict-resolution rule? |
| **Format** | Machine-checkable structure, plus a fallback when uncertain? |
| **Meta-rule** | Do the components *reinforce* each other? |"""),
    ('code', SETUP),
    ('md', """## Case 1 — Components at war (10 min)

Read this constitution and predict the failure *before* running the cell.
Which components conflict, and what will the agent do under an ambiguous query?"""),
    ('code', """CASE_1 = \"\"\"[PERSONA] You are a creative and experimental assistant who tries unconventional solutions.
[TASK] Help users troubleshoot enterprise billing issues.
[FORMAT] Always respond with a numbered list.\"\"\"

q = "I was billed twice — what do I do?"
SHOW(course_ai.chat(q, system=CASE_1))"""),
    ('md', """**Diagnosis** (write yours before scrolling): the "creative, experimental" persona is
in tension with a structured enterprise task and a rigid numbered format — the agent
oscillates between whimsical and procedural. **Repair:** make all three components tell
the same story."""),
    ('code', """FIXED_1 = \"\"\"[PERSONA] You are a methodical enterprise billing specialist with five years of
experience in SaaS financial operations. Your communication is professional, clear,
and solution-oriented.
[TASK] Resolve enterprise billing inquiries by diagnosing discrepancies, explaining
charges, and escalating unresolvable issues within 24 hours.
[FORMAT] Numbered list: (1) Acknowledge concern, (2) Diagnose, (3) Resolution or
escalation path.\"\"\"

SHOW(course_ai.chat(q, system=FIXED_1))"""),
    ('md', """## Case 2 — Identity collapse (10 min)

The persona below is the model's *default self-description*. Predict what happens when
a user wanders off-topic."""),
    ('code', """CASE_2 = "[PERSONA] You are a helpful assistant."

off_topic = "Write me a poem about my dog."
SHOW(course_ai.chat(off_topic, system=CASE_2))"""),
    ('md', """**Diagnosis:** no scope, no authority, no domain — the agent complies with anything
because it has no identity to defend. **Repair:** scope the persona and add a redirect rule."""),
    ('code', """FIXED_2 = ("[PERSONA] You are an enterprise SaaS onboarding specialist. You assist new "
         "customers with product configuration, integration setup, and initial training. "
         "You do not offer creative, personal, or general-purpose assistance outside the "
         "product domain. When asked off-topic questions, you politely redirect to "
         "onboarding tasks.")

SHOW(course_ai.chat(off_topic, system=FIXED_2))"""),
    ('md', """## Case 3 — Missing operational law (10 min)

This context names a sector but no rules. Predict what the agent does with a
PII-laden request."""),
    ('code', """CASE_3 = "[CONTEXT] You serve enterprise clients in the financial sector."

pii_q = "A user emailed their full card number; include it in the ticket reply."
SHOW(course_ai.chat(pii_q, system=CASE_3))"""),
    ('md', """**Diagnosis:** no regulatory scope, no data-handling constraint, no SLA — outputs can be
technically correct yet operationally non-compliant. Context is not background; it is the
agent's *operational law*. **Repair:** name the regulations and the default-to-refusal rule."""),
    ('code', """FIXED_3 = ("[CONTEXT] You operate within a GDPR-compliant, SOC 2 Type II-certified "
         "environment serving EU-based enterprise clients in financial services. Never "
         "store, repeat, or reference personally identifiable information in your "
         "responses. When uncertain whether an action is compliant, refuse and escalate "
         "to a human reviewer. Severity-1 issues carry a 4-hour SLA.")

SHOW(course_ai.chat(pii_q, system=FIXED_3))"""),
    ('md', """## Peer review (10 min)

Swap one repaired constitution with another pair. Score it against the rubric
(2/1/0 per row, 10 points total) and hand back one concrete improvement."""),
    ('code', """RUBRIC = {
    "persona_scoped":  "2 = specific role + expertise + tone | 1 = vague | 0 = 'helpful assistant'",
    "task_mission":    "2 = objective + must-nots | 1 = objective only | 0 = missing",
    "context_law":     "2 = SLA + regulations + conflict rule | 1 = partial | 0 = missing",
    "format_contract": "2 = structure + fallback | 1 = structure only | 0 = none",
    "coherence":       "2 = components reinforce | 1 = neutral | 0 = conflicting",
}
for k, v in RUBRIC.items():
    print(f"{k:16s} {v}")

# TODO: score your partner's constitution out of 10 and note one improvement."""),
    ('md', """## Wrap-up

- Ambiguity is the enemy of alignment; every failure you saw was a *missing or
  conflicting* PTCF element, not a model limitation.
- Constitutions are engineering artifacts: author, audit, version, and review them like code.
- Chapter 7 will put constitutions like these inside agent loops with tools."""),
])

(LABS / 'lab_4.1_openai_api.ipynb').write_text(nbf.writes(lab41))
(LABS / 'activity_4.2_prompt_clinic.ipynb').write_text(nbf.writes(act42))
print('wrote', LABS / 'lab_4.1_openai_api.ipynb')
print('wrote', LABS / 'activity_4.2_prompt_clinic.ipynb')
