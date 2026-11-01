"""One command that runs every check that does not need SolidWorks.

    python tests/check_all.py

Three layers, cheapest first:

    1. static_check.py   — structure: imported names, README sync, duplicate defs, property
                           traps, stale roadmap, wrapper signatures
    2. regressions.py    — one assertion per shipped bug (85 of them)
    3. compileall        — syntax (fails fast if a file is broken at all)

The behavioural layer (tests/smoke.py) needs a live SolidWorks and is NOT run here — it
takes ~6 minutes and depends on local-only model files.

Exit code 0 = every check passed.
"""
from __future__ import annotations

import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

STEPS = [
    ("syntax", [sys.executable, "-m", "compileall", "-q", os.path.join(ROOT, "src"),
                os.path.join(ROOT, "tests")]),
    ("structure", [sys.executable, os.path.join(HERE, "static_check.py")]),
    ("regressions", [sys.executable, os.path.join(HERE, "regressions.py")]),
]


def main():
    failures = []
    for name, cmd in STEPS:
        print("=" * 78)
        print(">> %s" % name)
        print("=" * 78)
        result = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True,
                                encoding="utf-8", errors="replace")
        output = (result.stdout or "").strip()
        if output:
            print(output)
        if result.returncode != 0:
            stderr = (result.stderr or "").strip()
            if stderr:
                print(stderr)
            failures.append(name)
        print()

    print("=" * 78)
    if failures:
        print("FAILED: %s" % ", ".join(failures))
        return 1
    print("all offline checks passed — for behaviour, run: python tests/smoke.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
