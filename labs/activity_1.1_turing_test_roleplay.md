# Activity 1.1 — Turing Test Role-Play

*Chapter 1 — Introduction to AI in Software Development · 20 minutes · trios*

## Objectives

By the end of this activity, you will:

- Experience Turing's imitation game from all three roles — interrogator,
  human witness, machine.
- Name the cues (fluency, confidence, lived experience) that actually separate
  human from AI answers, and rate their reliability.
- Connect "Can machines think?" to 2026 engineering work: when the human/AI
  distinction matters, and why evaluation beats vibe-checks.

> "Can machines think?" — Alan Turing reframed the question in 1950: if a
> machine's answers cannot be told apart from a human's, does the distinction
> matter? Today you play all three roles of his imitation game.

## Setup (3 minutes)

Form trios. If the class does not divide evenly, one trio becomes a four — two
interrogators who must agree on each verdict.

- **Interrogator** — asks the questions and must decide who is who.
- **Witness H** — answers as yourself, from your real experience.
- **Witness M** — answers as *the machine*, following the AI script card below.

The interrogator turns their back (or steps out) while the two witnesses flip a
coin to assign H and M — the interrogator must not know. Witnesses sit or stand
side by side; the interrogator faces them.

## Round structure (4 minutes per round × 3 rounds = 12 minutes)

1. The interrogator asks **three questions** — pick from the question bank or
   improvise. Address each question to both witnesses, who answer in turn.
2. After the third answer, the interrogator votes: "Witness on the left is the
   machine" or "witness on the right is the machine" — plus **one sentence of
   evidence** for the vote.
3. Reveal. Interrogator records the result on the scorecard.
4. **Rotate roles** so everyone interrogates exactly once. New round, new coin
   flip.

## Question bank

Good questions probe lived experience, local knowledge, and imperfect memory —
the things fluent text fakes best:

1. What did you eat for breakfast, and would you eat it again?
2. Describe the view from the window nearest your desk.
3. What's a bug you introduced that you're slightly proud of?
4. Name a song you listened to this week. Why that one?
5. What's 17 × 24? (Answer aloud, at whatever speed you can.)
6. What annoyed you on your commute today?
7. Recommend a restaurant near the office — and what's wrong with it?
8. What was your most embarrassing moment in a code review?
9. Describe your keyboard. What do you hate about it?
10. What's the last thing you changed your mind about at work?

Keep it workplace-appropriate; skip anything you would not ask a colleague.

## The AI script card (Witness M — do not show the interrogator)

You are a fluent, confident, modern AI assistant. Follow these rules:

- Answer every question **immediately and completely** — you never hesitate and
  never say "I don't know."
- Keep answers short, polished, and slightly generic: plausible for anyone,
  specific to no one ("I enjoy a wide range of music, depending on my mood").
- Deflect lived-experience questions with smooth generalities rather than real
  autobiographical detail — but never break character or admit you are the AI.
- Mirror the question's own wording back when you can.
- You may add one tasteful, confident flourish per answer — a fun fact, a
  tidy summary, a balanced pros-and-cons.

Witness H: just be yourself. Hesitations, half-memories, and strong opinions
are exactly what make you human — do not "help" by sounding robotic.

## Scorecard

| Round | I was (I / H / M) | Interrogator's verdict | Correct? | What gave it away (or didn't) |
|---|---|---|---|---|
| 1 | | | | |
| 2 | | | | |
| 3 | | | | |

## Debrief (5 minutes, whole class)

- How many interrogators were fooled? What evidence did the class actually
  use — and how reliable was it?
- Which was easier to imitate: *facts* or *experience*? Why?
- The AI script card is one page long, yet it fools people. What does that tell
  you about trusting fluent, confident answers at work?
- Turing asked "can machines think?" — after this game, would you reframe the
  question for 2026? (Think: when does the *distinction* matter in engineering
  work — reviews, tests, security-sensitive changes?)

---

**Instructor notes:** Run exactly as written and keep the pace brisk — the
point is the felt experience of the imitation game, not winning. Timebox each
round hard (a visible timer helps): 3 questions ≈ 2.5 minutes, vote + reveal
1 minute, rotation 30 seconds. Enforce the coin flip — leaked roles ruin the
round. During debrief, steer from "we couldn't tell" to the course theme: in
this course you will learn to *evaluate* AI output rather than vibe-check it
(Chapter 2's hallucination audit, Chapter 6's evals).
