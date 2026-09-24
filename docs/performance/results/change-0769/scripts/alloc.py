#!/usr/bin/env python3
"""Change 0769: allocation counts (record 0767's lane) per owner (counting-allocator probe).

One pinned `probe_alloc` process per probe case and arm runs one warmup and
five measured owners; every owner's allocated bytes, calls and peak live
bytes are kept, and the summary reports the first measured owner (all five
agree when an owner's allocations do not depend on process state).

Usage: alloc.py OUT_DIR > allocation.json
"""

import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run import CASES  # noqa: E402

CORE = "30"

ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0769"


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    summary = {}
    for name in CASES:
        if name.startswith("harness-"):
            continue
        for arm in ("base", "cand"):
            command = CASES[name](arm, 0, "/dev/null")
            command[0] = f"{ROOT}/bin/{arm}/probe_alloc"
            command[command.index("--samples") + 1] = "5"
            command[command.index("--warmups") + 1] = "1"
            completed = subprocess.run(["taskset", "-c", CORE] + command, capture_output=True, text=True)
            if completed.returncode != 0:
                raise SystemExit(f"{command} failed: {completed.stderr[-2000:]}")
            with open(f"{out_dir}/{name}-{arm}.json", "w") as handle:
                handle.write(completed.stdout)
            report = json.loads(completed.stdout)
            owners = report["allocation"]
            summary.setdefault(name, {})[arm] = {
                "owners": owners,
                "owners_identical": all(owner == owners[0] for owner in owners),
                "first": owners[0],
                "output_sha256": report["output_sha256"],
            }
    for name, arms in summary.items():
        base, cand = arms["base"]["first"], arms["cand"]["first"]
        arms["delta"] = {key: cand[key] - base[key] for key in base if isinstance(base[key], (int, float))}
        arms["same_output"] = arms["base"]["output_sha256"] == arms["cand"]["output_sha256"]
    json.dump(summary, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
