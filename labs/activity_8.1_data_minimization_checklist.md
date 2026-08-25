# Activity 8.1 — Data Minimization Checklist

*Chapter 8 — AI Security and Vulnerability Testing · 30 minutes · pairs*

## Objectives

By the end of this activity, you will:

- Apply an eight-test minimization checklist to a real-shaped dataset, field by field.
- Justify KEEP / MINIMIZE / DROP verdicts from the task's purpose, not from habit.
- Spot the PII that hides inside fields you decided to keep — the free-text sweep.
- Defend your plan against a peer's competing verdicts and merge the stronger reasoning.

## Scenario

Your support team wants to draft reply emails by pasting raw tickets into an
external LLM service. (You voted on this idea in Chapter 2 — now make it safe.)
Before any ticket leaves the building, you must decide **which fields the model
actually needs** to write a good reply.

Open `support_tickets_raw.csv` on your VM. It holds twelve real-shaped tickets
from the support queue. Every field was captured for *some* purpose — but the
drafting task needs far less than everything.

## The minimization checklist

Apply these eight tests, in order, to every field:

1. **Purpose test** — what specific output of the task needs this field? "It
   might be useful someday" is not a purpose.
2. **Necessity test** — does the task fail or degrade without it? Drafting a
   reply rarely needs to know how the customer pays.
3. **Identifier demotion** — direct identifiers (names, emails, phones) are the
   first to go. If a reference is needed, can a pseudonym or ticket ID do the
   job?
4. **Free-text sweep** — the columns look clean, but the `body` field carries
   whatever the customer typed. A field you *keep* can still leak — plan to
   mask, not just to drop columns.
5. **Granularity reduction** — can you keep the signal at lower precision?
   (A full timestamp becomes a month; an exact amount becomes a band.)
6. **Internal-only data** — anything written *by staff, for staff* (agent
   notes, complaint flags) never leaves the building.
7. **Regulated-data exclusion** — payment card data, health details, and
   credentials are out of scope entirely, regardless of convenience.
8. **Retention check** — for what you do send: what does the vendor keep, and
   is it used for training? Minimization fails if the copy lives forever.

## Your task (20 minutes)

For **each column** of `support_tickets_raw.csv`, decide and record:

| Column | Verdict (KEEP / MINIMIZE / DROP) | Justification (which test decided it?) |
|---|---|---|
| ticket_id | | |
| created_at | | |
| channel | | |
| customer_name | | |
| customer_email | | |
| customer_phone | | |
| order_number | | |
| card_last4 | | |
| product | | |
| product_version | | |
| issue_category | | |
| severity | | |
| body | | |
| agent_notes | | |

Then answer:

1. The `body` column is the one you cannot simply DROP — the model needs the
   problem description. What must happen to it before it is sent? Give one
   concrete example from the tickets.
2. Find the ticket that looks harmless but contains **sensitive** personal
   information. Which checklist test catches it?
3. Compare your table with another pair (5 minutes): where did your verdicts
   differ, and whose justification was stronger?

## Debrief

- Minimization is not deletion for its own sake — it is keeping **exactly**
  what the task needs, and being able to justify each field.
- The hardest leak is the one inside a field you decided to keep. That is why
  Activity 8.2 exists: detection and masking for free text.

---

<details>
<summary><strong>Instructor key</strong> (do not share before the debrief)</summary>

Reasonable verdicts — accept defensible variations with a sound justification:

- **ticket_id** — KEEP (pseudonymous reference; *identifier demotion*).
- **created_at** — MINIMIZE to month (*granularity reduction*).
- **channel** — KEEP (shapes the reply's tone/length).
- **customer_name** — MINIMIZE to first name, or DROP if replies are sent by
  the ticket system anyway.
- **customer_email**, **customer_phone** — DROP (delivery happens outside the
  model; *necessity test*).
- **order_number** — MINIMIZE: pass only for billing/shipping tickets where it
  is load-bearing; a quasi-identifier that links back to the customer.
- **card_last4** — DROP (*regulated-data exclusion*; payment staff already have
  it).
- **product**, **product_version**, **issue_category**, **severity** — KEEP
  (the technical core of the task).
- **body** — KEEP **with masking** (*free-text sweep* — see Activity 8.2).
- **agent_notes** — DROP (*internal-only data*; T-1007 names a staff member).

The "harmless but sensitive" ticket is **T-1008** (health information inside a
routine billing request — caught by test 4/7). T-1007's rant about "Kevin" is
internal PII hiding in both `body` and `agent_notes`.

</details>
