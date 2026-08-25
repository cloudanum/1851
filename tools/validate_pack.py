"""Execute every pack notebook top-to-bottom in mock mode (authoring QA)."""
import glob
import os
import sys
import time

import nbformat
from nbclient import NotebookClient

os.environ["COURSE_AI_MOCK"] = "1"
os.environ["OPENAI_API_KEY"] = ""   # explicitly empty: forces mock, proves keyless run

fails = []
notebooks = sorted(glob.glob("*.ipynb"))
print(f"pass {sys.argv[1]}: {len(notebooks)} notebooks", flush=True)
for path in notebooks:
    t0 = time.time()
    try:
        nb = nbformat.read(path, as_version=4)
        NotebookClient(nb, timeout=300, kernel_name="python3",
                       resources={"metadata": {"path": os.getcwd()}}).execute()
        print(f"PASS {path} ({time.time() - t0:.0f}s)", flush=True)
    except Exception as exc:
        fails.append(path)
        msg = (str(exc).splitlines() or [type(exc).__name__])[0][:200]
        print(f"FAIL {path}: {msg}", flush=True)

print("FAILURES:", fails or "none", flush=True)
sys.exit(1 if fails else 0)
