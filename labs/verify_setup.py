"""verify_setup.py — 60-second smoke test for the 1851 lab environment.

Run from the pack root after `pip install -r requirements.txt`:

    python verify_setup.py

Checks the lab dependencies import and that course_ai answers in whichever
mode the machine is in (mock when keyless / COURSE_AI_MOCK=1, live with a key).
Exits non-zero if anything fails.
"""
import sys

failures = []


def check(label, fn):
    try:
        fn()
        print(f"  ok    {label}")
    except Exception as exc:  # noqa: BLE001 — a smoke test reports everything
        failures.append(label)
        print(f"  FAIL  {label}: {exc}")


def _deps():
    import numpy, pandas, pytest, sklearn  # noqa: F401


def _sdk():
    import openai  # noqa: F401


def _chat_roundtrip():
    import course_ai
    out = course_ai.chat("Reply with the single word: pong")
    assert isinstance(out, str) and out.strip(), "empty chat() reply"


def _json_roundtrip():
    import course_ai
    schema = {"type": "object",
              "properties": {"entities": {"type": "array"}},
              "required": ["entities"]}
    out = course_ai.chat_json(
        "You are a PII detection engine.\n\nTICKET T-1009:\n"
        "The upgrade guide is unclear about the broker restart step.", schema)
    assert "entities" in out and isinstance(out["entities"], list)


def _ml_stack():
    import numpy as np
    import pandas as pd
    from sklearn.tree import DecisionTreeClassifier
    df = pd.DataFrame({"x": np.arange(8), "y": [0, 0, 0, 0, 1, 1, 1, 1]})
    tree = DecisionTreeClassifier(max_depth=1, random_state=0).fit(df[["x"]], df["y"])
    assert tree.predict(pd.DataFrame({"x": [7]}))[0] == 1


print("1851 lab environment check")
import course_ai  # noqa: E402
print("course_ai mode:", course_ai.mode(), "\n")
check("lab dependencies (numpy, pandas, pytest, sklearn)", _deps)
check("openai SDK import", _sdk)
check("course_ai.chat round-trip", _chat_roundtrip)
check("course_ai.chat_json round-trip", _json_roundtrip)
check("pandas + scikit-learn train/predict", _ml_stack)

if failures:
    print(f"\n{len(failures)} check(s) FAILED — fix the environment before class.")
    sys.exit(1)
print("\nAll checks passed — the environment is ready.")
