#!/usr/bin/env python3
"""Change 0745 allocation lane.

Runs the counting-allocator probe (`probe_alloc`) once per case and arm with
two warmups and three measured owners; each owner's region records allocated
and deallocated bytes, allocation calls, peak live bytes relative to entry and
bytes still live at exit (the owner's retained output). Regions must agree
exactly across the three owners of a process.

Usage: alloc.py OUT_DIR > allocation.json
"""

import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run import ARMS, CASES, CORE, ROOT  # noqa: E402

PROBE_CASES = [name for name in CASES if not name.startswith(("harness-", "p0734-"))]


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    result = {}
    for name in PROBE_CASES:
        for arm in ARMS:
            command = CASES[name](arm, 0, "/dev/null")
            command[0] = f"{ROOT}/bin/{arm}/probe_alloc"
            for flag, value in (("--warmups", "2"), ("--samples", "3")):
                command[command.index(flag) + 1] = value
            completed = subprocess.run(["taskset", "-c", CORE] + command, capture_output=True, text=True, check=True)
            with open(f"{out_dir}/{name}-{arm}.json", "w") as handle:
                handle.write(completed.stdout)
            regions = json.loads(completed.stdout)["allocation"]
            if any(region != regions[0] for region in regions):
                raise SystemExit(f"{name} {arm}: allocation regions differ between owners")
            result.setdefault(name, {})[arm] = regions[0]
    json.dump(result, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
