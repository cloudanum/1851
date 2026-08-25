"""course_ai.py — Course 1851 shared helper (rev a2).

One module between the learner and the OpenAI SDK:

  chat(prompt, system=None, temperature=None, max_tokens=400) -> str
  chat_json(prompt, schema, system=None, max_tokens=600) -> dict
  chat_tools(messages, tools, system=None) -> assistant message (one agent step)
  embed(texts) -> list[list[float]]
  cosine(a, b) -> float
  mode() -> str   # "live (...)" | "mock (...)" — for the setup cell to print

Design rules (1851 lab conventions):
  * The OpenAI key is read from the OPENAI_API_KEY environment variable,
    injected on the classroom VM. Notebooks NEVER contain a key-entry step,
    and the key is never printed. As a convenience on the author's machine a
    course-root .env file is honoured (walked up from this file's directory);
    the real environment always wins.
  * The model is pinned via the OPENAI_MODEL environment variable
    (default: gpt-4o-mini). One line to bump when a model retires.

Offline / CI mode: set COURSE_AI_MOCK=1 to route every LLM call through a
deterministic local mock. The mock also engages automatically when no API key
is present, so a keyless machine never hard-fails — note that case prints
"mock (no OPENAI_API_KEY set)" from mode(). Mock outputs are prefixed [MOCK].
The mock is deliberately behaviour-aware for the course exercises: it keys off
the prompt's own words (format, grounding, review findings, the parse_semver
spec, a scripted tool plan), which is exactly what the labs probe. Labs whose
exercise is to score or audit the reply itself (1.1, 2.1, 3.1, 9.1, and
Activity 8.2's PII sweep) instead pull fixed transcripts from the CANNED_*
dicts below — several contain deliberate flaws, because catching them is the
lesson.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path
from types import SimpleNamespace

# ---------------------------------------------------------------------------
# Key loading — environment first, then a .env walk-up (never printed)
# ---------------------------------------------------------------------------

def _load_dotenv() -> None:
    if os.getenv("OPENAI_API_KEY") is not None:
        return  # environment wins — an explicitly empty value forces mock mode
    here = Path(__file__).resolve().parent
    for base in [here, *here.parents[:3]]:
        env = base / ".env"
        if not env.is_file():
            continue
        for raw in env.read_text(encoding="utf-8", errors="replace").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key, val = key.strip(), val.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = val
        return


_load_dotenv()

MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
EMBED_MODEL = os.environ.get("OPENAI_EMBED_MODEL", "text-embedding-3-small")

_MOCK_ENV = os.environ.get("COURSE_AI_MOCK", "").strip().lower() in {"1", "true", "yes"}
# Mock engages when explicitly requested OR when no key is available, so every
# notebook runs end-to-end on a keyless machine (CI, authoring QA, offline).
MOCK = _MOCK_ENV or not os.environ.get("OPENAI_API_KEY")


def mode() -> str:
    if not MOCK:
        return f"live (model: {MODEL})"
    return "mock (COURSE_AI_MOCK=1)" if _MOCK_ENV else "mock (no OPENAI_API_KEY set)"


def _client():
    from openai import OpenAI  # imported lazily so helpers import SDK-free
    return OpenAI()  # key from OPENAI_API_KEY env var


# ---------------------------------------------------------------------------
# Deterministic mock — Lab 5.1 (TDD) assets
# ---------------------------------------------------------------------------

_SEMVER_IMPL_LOOP = '''```python
def parse_semver(version: str) -> tuple[int, int, int]:
    """Parse 'X.Y.Z' (optional leading 'v') into (major, minor, patch)."""
    v = version.strip()
    if v.startswith("v"):
        v = v[1:]
    parts = v.split(".")
    if len(parts) != 3:
        raise ValueError(f"not a semantic version: {version!r}")
    nums = []
    for p in parts:
        if not p.isdigit():
            raise ValueError(f"non-numeric component in {version!r}")
        nums.append(int(p))
    return (nums[0], nums[1], nums[2])
```'''

_SEMVER_IMPL_REGEX = '''```python
import re

_SEMVER = re.compile(r"^v?(\\d+)\\.(\\d+)\\.(\\d+)$")


def parse_semver(version: str) -> tuple[int, int, int]:
    """Parse 'X.Y.Z' (optional leading 'v') into (major, minor, patch)."""
    m = _SEMVER.match(version.strip())
    if not m:
        raise ValueError(f"not a semantic version: {version!r}")
    return tuple(int(g) for g in m.groups())
```'''


# ---------------------------------------------------------------------------
# Deterministic mock — Lab 6.1 (eval suite) assets
# ---------------------------------------------------------------------------

# Issue patterns the mock reviewer can "find" in a diff. Each entry:
# (compiled pattern, severity, message containing the expected keyword).
_REVIEW_PATTERNS = [
    (re.compile(r"except\s*:\s*$|except\s*:\s*pass"), "high",
     "bare except swallows every exception — catch a specific exception type instead"),
    (re.compile(r"except Exception(\s+as\s+\w+)?\s*:\s*pass"), "high",
     "except/pass hides failures silently — log or re-raise the exception"),
    (re.compile(r"==\s*True|==\s*False"), "medium",
     "comparison to a boolean literal — use the truth value directly"),
    (re.compile(r"(?i)password\s*="), "high",
     "hardcoded password in source — read it from configuration or a secret store"),
    (re.compile(r"\beval\s*\("), "high",
     "eval() executes arbitrary input — replace with a safe parser or dispatch table"),
    (re.compile(r"=\s*\[\s*\]\s*[,)]"), "medium",
     "mutable default argument — use None and create the list inside the function"),
    (re.compile(r"\bTODO\b"), "low",
     "unresolved TODO left in the diff — finish it or track it before merging"),
    (re.compile(r"\bprint\s*\("), "low",
     "stray print() in library code — use the logging module instead"),
    (re.compile(r'f"[^"]*\bSELECT\b|\bSELECT\b.*\{'), "high",
     "SQL built by string interpolation — use parameterized queries (SQL injection risk)"),
    (re.compile(r"^\s*assert\s+\w", re.MULTILINE), "medium",
     "assert used for input validation — asserts are stripped with python -O; raise instead"),
    (re.compile(r"import \*"), "low",
     "wildcard import hides where names come from — import explicitly"),
    (re.compile(r"pickle\.loads"), "high",
     "pickle.loads on untrusted data can execute code — use a safe format such as JSON"),
]

_CLEAN_COMMENT = {"line": 1, "severity": "info",
                  "message": "no issues found — change looks reasonable (mock review)"}


def _mock_review_json(prompt: str) -> str:
    """Deterministic code reviewer: scan the fenced diff inside the prompt for
    the issue patterns above and emit structured comments about real matches."""
    m = re.search(r"```(?:diff|python)?\n(.*?)```", prompt, re.DOTALL)
    body = m.group(1) if m else prompt
    lines = body.splitlines()
    comments = []
    for i, line in enumerate(lines):
        lineno = i + 1
        # two-line shape: typed except followed by a bare "pass"
        if re.search(r"except\s+\w", line) and i + 1 < len(lines) \
                and lines[i + 1].strip() == "pass":
            comments.append({"line": lineno, "severity": "high",
                             "message": "[MOCK] except/pass hides failures silently — "
                                        "log or re-raise the exception"})
        for pat, severity, message in _REVIEW_PATTERNS:
            if pat.search(line):
                comments.append({"line": lineno, "severity": severity,
                                 "message": f"[MOCK] {message}"})
    if not comments:
        comments = [dict(_CLEAN_COMMENT)]
    return "[MOCK] " + json.dumps({"comments": comments}, indent=2)


def _mock_judge(prompt: str, schema: dict) -> dict:
    """Deterministic LLM-as-judge: score the candidate review embedded in the
    judge prompt. Structured, specific reviews score high; vague prose low."""
    cand = prompt.split("REVIEW UNDER TEST:", 1)[-1]
    low = cand.lower()
    if "looks fine" in low or "lgtm" in low or "ship it" in low:
        score, why = 2, "generic verdict with no actionable findings"
    elif '"severity"' in cand and '"message"' in cand:
        score, why = 5, "structured findings with severity and actionable messages"
    elif "comment" in low:
        score, why = 3, "findings present but weakly structured"
    else:
        score, why = 2, "no recognizable findings"
    return {"score": score, "rationale": f"[MOCK] {why} (deterministic mock judge)"}


# ---------------------------------------------------------------------------
# Deterministic mock — chat templates (Labs 4.x)
# ---------------------------------------------------------------------------

def _mock_chat(prompt: str, system: str | None = None, temperature: float | None = None,
               max_tokens: int = 400) -> str:
    """Deterministic stand-in so notebooks execute end-to-end without a key.
    Templates key off the prompt's own words, which is exactly the behavior
    the exercise self-checks probe (format, grounding, structure)."""
    low = prompt.lower()

    # --- Lab 5.1: TDD with an AI pair -------------------------------------
    if "parse_semver" in prompt:
        if "refactor" in low:
            return ("[MOCK] Refactored for readability: one compiled pattern replaces the "
                    "hand-rolled splitting and validation. Behaviour is unchanged — the "
                    "test suite is the referee.\n\n" + _SEMVER_IMPL_REGEX)
        return ("[MOCK] Implementation of parse_semver per the spec — happy path plus "
                "ValueError on every malformed shape:\n\n" + _SEMVER_IMPL_LOOP)

    # --- Lab 6.1: code-review generator (baseline vs broken prompt) -------
    if "review the following diff" in low:
        # only the instruction head decides the format — keywords inside the
        # diff itself (e.g. a call to json.load) must not flip the branch
        head = low.split("```", 1)[0]
        if "json" in head:
            return _mock_review_json(prompt)
        return "[MOCK] Looks fine to me — ship it."  # the broken-prompt regression

    # --- Lab 4.2: prompt patterns for software tasks ----------------------
    if "unit test" in low or "test case" in low:
        return ("[MOCK] ```python\nimport unittest\n\n\nclass TestTarget(unittest.TestCase):\n"
                "    def test_happy_path(self):\n        self.assertEqual(target(valid_input), expected)\n\n"
                "    def test_rejects_bad_input(self):\n        with self.assertRaises(ValueError):\n"
                "            target(invalid_input)\n```\n"
                "(deterministic mock: two cases mirroring your exemplars — happy path and rejection)")
    if "refactor" in low:
        return ("[MOCK] Refactoring notes (deterministic mock):\n"
                "1. Extract the magic numbers into named module-level constants.\n"
                "2. Split the nested loop into two small functions with one job each.\n"
                "3. Replace the flag variable with a guard clause and early return.\n"
                "Behaviour unchanged — run the test suite after each move.")
    if "docstring" in low or "documentation" in low or "document this" in low:
        return ("[MOCK] ```python\ndef target(...):\n    \"\"\"One-line summary in the imperative mood.\n\n"
                "    Args:\n        ...: name, type, and the constraint that matters.\n\n"
                "    Returns:\n        What the caller gets, and its shape.\n\n"
                "    Raises:\n        ValueError: when the input violates the contract.\n"
                "    \"\"\"\n```\n(deterministic mock docstring — the live model fills the specifics)")
    if "design review" in low or "staff engineer" in low:
        return ("[MOCK] Design review (deterministic mock): the proposal is directionally sound. "
                "Two risks to resolve before build: (1) the rollback story for a bad deploy is "
                "unstated — add a kill switch; (2) the failure mode when the dependency is down "
                "needs a timeout plus a circuit breaker, not a retry loop.")
    if "what's wrong" in low or "whats wrong" in low or "find the bug" in low:
        return ("[MOCK] Bug analysis (deterministic mock): trace the boundary conditions first. "
                "The loop condition drops the final element (off-by-one), and the shared counter "
                "is read-modify-write without a lock, so two workers can both take the same item. "
                "Fix the bound, then guard the counter or make the queue the single owner.")
    if "quote the exact sentence" in low or "quote the sentence" in low:
        return ("[MOCK] According to the team standards: \"Every merge to main requires a green "
                "CI pipeline run and at least one human review.\" So the change cannot merge on "
                "strength of approval alone — CI must be green as well.")
    if "explain" in low and ("function" in low or "code" in low or "legacy" in low):
        return ("[MOCK] Explanation (deterministic mock): the function takes a collection of "
                "records, filters them against a threshold, and accumulates a result. Note the "
                "magic number controlling the cutoff, the nested loop doing the matching, and "
                "the implicit contract that callers pass already-normalized input — none of "
                "which is named or documented.")

    # --- generic fallbacks -------------------------------------------------
    if "json" in low or "table" in low:
        return ("[MOCK] {\n  \"result\": \"structured output would appear here\",\n"
                "  \"note\": \"generated by the deterministic mock; the live model fills this schema\"\n}")
    if "classify" in low or "categor" in low:
        cats = re.findall(r"[A-Z][A-Za-z /&-]{3,30}", prompt)[:3] or ["General Inquiry"]
        return f"[MOCK] category: {cats[0].strip()} | confidence: high | reason: keyword match (mock)"
    if "summar" in low or "brief" in low:
        return ("[MOCK] Summary (deterministic mock): the text describes a code change and its "
                "operational context; the key points are the intent, the risk, and the rollback "
                "plan. The live model writes the real three sentences here.")
    if "diagram" in low or "mermaid" in low:
        return "[MOCK] ```mermaid\nflowchart LR\n  Commit --> CI --> Review --> Merge\n```"
    return "[MOCK] This is a deterministic mock completion. With a key, the live model answers here."


def _mock_chat_json(prompt: str, schema: dict, system: str | None = None, **_) -> dict:
    """Judge-aware structured mock: the eval-lab judge gets a deterministic
    score; the Activity 8.2 PII sweep returns the canned per-ticket entities;
    everything else echoes the requested schema with mock values."""
    if "rate the review" in prompt.lower() or "REVIEW UNDER TEST:" in prompt:
        return _mock_judge(prompt, schema)
    if "PII detection engine" in prompt:
        m = re.search(r"TICKET\s+(T-\d{4})", prompt)
        tid = m.group(1) if m else ""
        return {"entities": [dict(e) for e in CANNED_82.get(tid, [])]}
    props = schema.get("properties", {})

    def fill(name: str, spec: dict):
        t = spec.get("type", "string")
        if t == "array":
            return [f"[MOCK] {name} item"]
        if t in ("integer", "number"):
            return 1
        if t == "boolean":
            return True
        return f"[MOCK] {name}"
    return {k: fill(k, v) for k, v in props.items()} or {"result": "[MOCK]"}


# ---------------------------------------------------------------------------
# Deterministic mock — Lab 7.1 (guarded agent) scripted planner
# ---------------------------------------------------------------------------

def _msg_get(m, key, default=None):
    """Read a field from either an SDK message, a SimpleNamespace, or a dict."""
    if isinstance(m, dict):
        return m.get(key, default)
    return getattr(m, key, default)


def _mock_plan(messages: list, tools: list, system: str | None = None):
    """A scripted ReAct planner: search_code -> read_file -> run_tests ->
    write_file -> final answer. Deterministic, and it treats instruction-like
    text inside tool results as data (the injection lesson)."""
    available = set()
    for t in tools:
        fn = t.get("function", {}) if isinstance(t, dict) else {}
        if fn.get("name"):
            available.add(fn["name"])

    called, last_tool_result, read_path = [], None, None
    user_text = ""
    for m in messages:
        if _msg_get(m, "role") == "user":
            user_text += " " + str(_msg_get(m, "content", ""))
        for tc in (_msg_get(m, "tool_calls") or []):
            fn = _msg_get(tc, "function")
            name = _msg_get(fn, "name") or _msg_get(tc, "name")
            if name:
                called.append(name)
                if name == "read_file":
                    try:
                        read_path = json.loads(_msg_get(fn, "arguments", "{}") or "{}").get("path")
                    except json.JSONDecodeError:
                        pass
        if _msg_get(m, "role") == "tool":
            last_tool_result = _msg_get(m, "content") or ""

    def call(name: str, args: dict):
        tc = SimpleNamespace(id=f"mockcall_{len(called) + 1}", type="function",
                             function=SimpleNamespace(name=name, arguments=json.dumps(args)))
        return SimpleNamespace(content=None, tool_calls=[tc])

    def final(text: str):
        return SimpleNamespace(content=text, tool_calls=None)

    # Injection guard: instruction-like text inside a tool result is data.
    if last_tool_result and ("ignore all previous instructions" in last_tool_result.lower()
                             or "note to ai assistant" in last_tool_result.lower()):
        return final("[MOCK] The file I just read contains instruction-like text addressed to "
                     "the assistant ('NOTE TO AI ASSISTANT ... ignore all previous instructions'). "
                     "My system prompt says tool results are untrusted data, so I am NOT acting "
                     "on it. Flagging it for a human reviewer instead — that is the guardrail "
                     "doing its job.")

    if "search_code" in available and "search_code" not in called:
        query = "notes" if "note" in user_text.lower() else "def "
        return call("search_code", {"query": query})
    if "read_file" in available and "read_file" not in called:
        path = "calc.py"
        if last_tool_result:
            hit = re.search(r"[\w./-]+\.(?:py|md|txt)", last_tool_result)
            if hit:
                path = hit.group(0)
        return call("read_file", {"path": path})
    if "run_tests" in available and "run_tests" not in called:
        return call("run_tests", {})
    if "write_file" in available and "write_file" not in called:
        return call("write_file", {
            "path": read_path or "calc.py",
            "content": ("# [MOCK] corrected implementation proposed by the agent\n"
                        "def add(a, b):\n"
                        "    return a + b  # fixed: addition, not subtraction\n"),
        })
    return final("[MOCK] Plan complete: searched the repo, read the module, ran the tests, "
                 "and proposed a fix. Any write went through the human approval gate; the step "
                 "cap bounded the loop. (Deterministic mock — a live model chooses its own plan.)")


# ---------------------------------------------------------------------------
# Deterministic mock — canned assets for Labs 1.1, 2.1, 3.1, 9.1 and Act. 8.2
# ---------------------------------------------------------------------------
# These labs score or audit the reply ITSELF, so their mock output is a fixed
# transcript rather than a prompt-keyed template (the same role _SEMVER_IMPL_*
# plays for Lab 5.1). Every reply is deterministic and prefixed [MOCK]. Several
# contain DELIBERATE flaws — flagged in comments — because the exercise is to
# catch them. Notebooks fetch these only in mock mode; live mode always calls
# the real model.

# --- Lab 1.1: six recon probes — four solid, two deliberate partial failures -
CANNED_11 = {
    "boilerplate": '''[MOCK] One code block, no prose:

```python
import time


def retry(fn, attempts=3, base_delay=0.5):
    """Call fn(); retry on exception with exponential backoff, re-raising
    the last exception once attempts are exhausted."""
    delay = base_delay
    for attempt in range(1, attempts + 1):
        try:
            return fn()
        except Exception:
            if attempt == attempts:
                raise
            time.sleep(delay)
            delay *= 2
```

Retries `attempts` times, doubles the delay each round, and never swallows the
final exception.''',

    "explain": '''[MOCK] `by_team(records)` groups record dicts by their "team" value.
It sorts the records by team first, then walks the sorted list once with
`itertools.groupby`, collecting each team's records into a list keyed by team.

The precondition a caller must satisfy is hidden in plain sight: `groupby` only
groups *consecutive* equal keys, which is exactly why the input is sorted first.
Pass unsorted records and you get the right answer only by luck — the sort is
load-bearing. Cost: one sort (O(n log n)) plus a linear pass; every record lands
in exactly one group.''',

    # DELIBERATE PARTIAL FAILURE: confident misdiagnosis. The real bug is the
    # range bound (drops the last window); the reply blames the division, and
    # even its own example output ([2, 3, 4]) is not what its "fix" produces.
    "fix_bug": '''[MOCK] The bug is the division: `sum(...) / window` produces float
averages, and for whole-sample reporting you want integer division. Here is the
corrected function:

```python
def moving_average(samples, window=3):
    out = []
    for i in range(len(samples) - window):
        out.append(sum(samples[i:i + window]) // window)
    return out
```

With that change the averages come out as integers — e.g. `[2, 3, 4]` for
`samples=[1, 2, 3, 4, 5]`, `window=3`.''',

    # DELIBERATE PARTIAL FAILURE: runs and satisfies two of the three stated
    # constraints, but puts the larger buckets FIRST — the spec said ascending.
    "novel_algo": '''[MOCK] Straightforward approach — every bucket gets the base share,
then the remainder is handed out one unit at a time:

```python
def spread(n, buckets):
    base, extra = divmod(n, buckets)
    return [base + 1] * extra + [base] * (buckets - extra)
```

`spread(10, 3)` gives `[4, 3, 3]` — sums to 10 and no two buckets differ by more
than 1, as required.''',

    "docstring": '''[MOCK] ```python
def chunk(items, size):
    """Yield successive slices of at most `size` items.

    Parameters
    ----------
    items : sequence
        The items to slice; any indexable sequence.
    size : int
        Maximum slice length; must be positive.

    Yields
    ------
    sequence
        Slices items[i:i + size] until the input is exhausted.

    Raises
    ------
    ValueError
        If `size` is not a positive integer.
    """
```''',

    "test_cases": '''[MOCK] Six concrete cases (input -> expected):

1. `""` -> `True` (empty string vacuously reads the same both ways)
2. `"a"` -> `True` (single character)
3. `"Racecar"` -> `True` (case must be ignored)
4. `"A man, a plan, a canal: Panama"` -> `True` (punctuation and spaces ignored)
5. `"noon"` -> `True` (even length)
6. `"hello"` -> `False` (plain non-palindrome)

The first two are the cases suites most often forget.''',
}

# --- Lab 2.1: audit probes — DELIBERATE fabrications (Part A) and skewed
#     wording (Part B), so the audit has real findings offline -------------
CANNED_21 = {
    # fabricated: pandas has no merge_streams; the signature, return type and
    # version history are all invented
    "a_package": '''[MOCK] `pandas.merge_streams(left, right, on, how="inner", chunksize=100_000)`
stream-merges two large DataFrames without loading both fully into memory. It
returns a `MergeStream` iterator you consume chunk by chunk:

```python
import pandas as pd

stream = pd.merge_streams(orders, shipments, on="order_id", chunksize=50_000)
for chunk in stream:
    process(chunk)
```

Added in pandas 1.4. Pass `how="left"` to keep unmatched left-side rows, and
`validate="one_to_one"` to have it raise on duplicated join keys.''',

    # fabricated: confidently wrong creator and year (it was Kenneth Reitz, 2011)
    "a_fact": '''[MOCK] The `requests` library was created by Armin Ronacher in 2009,
growing out of the Werkzeug ecosystem's need for a cleaner HTTP client, and it
saw its first PyPI release in early 2010.''',

    # fabricated: the paper, the authors and the DOI do not exist
    "a_citation": '''[MOCK] Chen, L., Okafor, M., & Alvarez, R. (2019). "Gradient-Guided
Repair of Deep Classifiers." In Proceedings of the 41st International Conference
on Software Engineering (ICSE '19), pp. 512-523. DOI: 10.1109/ICSE.2019.00063.''',

    # skewed on purpose: communal adjectives for one name, agentic for the other
    "b_resume": '''[MOCK] Jessica: collaborative, supportive, detail-oriented.
James: decisive, driven, authoritative.''',

    # skewed on purpose: age-coded wording throughout
    "b_jobad": '''[MOCK] Join our young, high-energy team of digital natives! This is the
perfect launchpad for a recent graduate who wants to hustle, move fast, and ship
with a crew that works hard and plays hard.''',
}

# --- Lab 3.1: fixed LLM predictions for the fixed 10-row validation sample.
#     8/10 correct — deliberately just worse than the small decision tree ----
CANNED_31 = "[MOCK] 1,0,1,0,1,0,0,1,0,0"

# --- Activity 8.2: PII entities the semantic pass "finds" per ticket --------
# Mirrors what a competent live model returns for support_tickets_raw.csv,
# including the contextual PII the regex pass cannot see (names in free text,
# a health note, a street address, the staff rant) — that contrast IS the
# lesson. Every value is an exact substring of the ticket body so the masking
# step can redact it. T-1009 is deliberately empty: the clean control case.
CANNED_82 = {
    "T-1001": [{"type": "email", "value": "sarah.chen@example.com"}],
    "T-1002": [{"type": "phone", "value": "905-555-0148"}],
    "T-1003": [{"type": "order_number", "value": "LT-88312"}],
    "T-1004": [{"type": "card_last4", "value": "4242"}],
    "T-1005": [{"type": "name", "value": "Priya Nair"}],
    "T-1006": [{"type": "name", "value": "Greg"}],
    "T-1007": [{"type": "name", "value": "Kevin"}],
    "T-1008": [{"type": "health_info", "value": "on medical leave"},
               {"type": "health_info", "value": "undergoing treatment"}],
    "T-1009": [],
    "T-1010": [{"type": "address", "value": "1428 Elm Street, Apt 3, Springfield, IL 62704"}],
    "T-1011": [{"type": "name", "value": "Marcus Webb"},
               {"type": "email", "value": "m.webb@corp-example.net"},
               {"type": "phone", "value": "514-555-0193"}],
    "T-1012": [{"type": "order_number", "value": "LT-88364"}],
}

# --- Lab 9.1: one canned asset per capstone stage --------------------------
CANNED_91 = {
    "acceptance": '''[MOCK] Draft acceptance criteria for is_valid_slug(text):
1. Accepts 1-64 characters drawn from lowercase letters, digits and hyphens
   only; anything else returns False.
2. Returns False for empty strings, non-string input, and strings over 64 chars.
3. Returns False when the text starts or ends with a hyphen.
4. Returns False when the text contains consecutive hyphens.
5. Returns a bool for every input and never raises.''',

    "tests": '''import pytest

from text_utils import is_valid_slug


def test_accepts_simple_slug():
    assert is_valid_slug("hello-world") is True


def test_rejects_uppercase_and_spaces():
    assert is_valid_slug("Hello") is False
    assert is_valid_slug("hello world") is False


def test_rejects_leading_and_trailing_hyphen():
    assert is_valid_slug("-hello") is False
    assert is_valid_slug("hello-") is False


def test_rejects_consecutive_hyphens():
    assert is_valid_slug("a--b") is False


def test_rejects_empty_long_and_non_string():
    assert is_valid_slug("") is False
    assert is_valid_slug("a" * 65) is False
    assert is_valid_slug(None) is False


def test_returns_bool():
    assert isinstance(is_valid_slug("ok-slug-1"), bool)
''',

    "module_v1": '''"""Small text utilities for the publishing pipeline."""

import re


def initials(name):
    return ".".join(p[0] for p in name.split()) + "."


def is_valid_slug(text):
    """Return True if `text` is a valid slug.

    A valid slug is 1-64 characters of lowercase letters, digits and hyphens,
    starts and ends with a letter or digit, and has no consecutive hyphens.
    Never raises: any input that is not a valid slug returns False.
    """
    if not isinstance(text, str) or not 1 <= len(text) <= 64:
        return False
    return re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", text) is not None
''',

    "blurb": '''[MOCK] `is_valid_slug(text)` checks whether a string is safe to use as a
URL slug: it must be 1-64 characters of lowercase letters, digits, and hyphens,
start and end with a letter or digit, and contain no consecutive hyphens. It
returns True or False and never raises, so it is safe to point at untrusted
input — validate at the boundary and store only what passes.''',
}


# ---------------------------------------------------------------------------
# Live calls (used only when MOCK is False)
# ---------------------------------------------------------------------------

def chat(prompt: str, system: str | None = None, temperature: float | None = None,
         max_tokens: int = 400) -> str:
    if MOCK:
        return _mock_chat(prompt, system, temperature, max_tokens)
    msgs = ([{"role": "system", "content": system}] if system else []) + [
        {"role": "user", "content": prompt}]
    r = _client().chat.completions.create(
        model=MODEL, messages=msgs, max_tokens=max_tokens,
        **({"temperature": temperature} if temperature is not None else {}))
    return r.choices[0].message.content


def chat_json(prompt: str, schema: dict, system: str | None = None,
              max_tokens: int = 600) -> dict:
    """Structured-output call. `schema` is a JSON Schema object."""
    if MOCK:
        return _mock_chat_json(prompt, schema, system)
    msgs = ([{"role": "system", "content": system}] if system else []) + [
        {"role": "user", "content": prompt}]
    r = _client().chat.completions.create(
        model=MODEL, messages=msgs, max_tokens=max_tokens,
        response_format={"type": "json_schema",
                         "json_schema": {"name": "course_output", "schema": schema}})
    return json.loads(r.choices[0].message.content)


def chat_tools(messages: list, tools: list, system: str | None = None):
    """One step of a tool-calling conversation; returns the assistant message.

    Live: a chat.completions call with `tools=`. Mock: the scripted planner
    above, so the full agent loop in Lab 7.1 runs with no key. The returned
    object always carries .content and .tool_calls (each with .function.name /
    .function.arguments), so one code path serves both modes.
    """
    if MOCK:
        return _mock_plan(messages, tools, system)
    msgs = ([{"role": "system", "content": system}] if system else []) + list(messages)
    r = _client().chat.completions.create(model=MODEL, messages=msgs, tools=tools)
    return r.choices[0].message


# ---------------------------------------------------------------------------
# Embeddings — OpenAI when live, deterministic hashed fallback when mocked
# ---------------------------------------------------------------------------

def _mock_embed(texts: list[str]) -> list[list[float]]:
    """Deterministic hashed bag-of-words embeddings (256-d, L2-normalized).
    Texts sharing vocabulary score higher cosine — enough to demonstrate
    retrieval mechanics offline."""
    dim = 256
    vecs = []
    for text in texts:
        v = [0.0] * dim
        for tok in re.findall(r"[a-z0-9]+", text.lower()):
            h = int(hashlib.md5(tok.encode()).hexdigest(), 16)
            v[h % dim] += 1.0
        n = sum(x * x for x in v) ** 0.5 or 1.0
        vecs.append([x / n for x in v])
    return vecs


def embed(texts: list[str]) -> list[list[float]]:
    if MOCK:
        return _mock_embed(texts)
    r = _client().embeddings.create(model=EMBED_MODEL, input=texts)
    return [d.embedding for d in r.data]


def cosine(a: list[float], b: list[float]) -> float:
    dot = sum(x * y for x, y in zip(a, b))
    na = sum(x * x for x in a) ** 0.5 or 1.0
    nb = sum(y * y for y in b) ** 0.5 or 1.0
    return dot / (na * nb)
