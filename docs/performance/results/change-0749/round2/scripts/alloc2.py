#!/usr/bin/env python3
"""Change 0749 round two: counting-allocator lane, arms A, B and C.

One pinned `probe_alloc` process per probe case and arm, 2 warmups and 3
measured owners (1 warmup and 2 owners for the generated files). Each
owner's allocation region is kept; the summary reports the per-owner values,
which are identical across owners of one process, and flags any case where
they are not.

Usage: alloc2.py OUT_DIR > allocation.json
"""

import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run2 import CASES, CORE, binary  # noqa: E402

FIELDS = ("allocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes")


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    summary = {}
    for name, case in CASES.items():
        if case["arms"] != "probe":
            continue
        generated = "-v3-" in name or "-v4-" in name
        for arm in ("A", "B", "C"):
            command = case["build"](arm, 0, "/dev/null")
            command[0] = binary(arm, "probe_alloc", 0)
            command[command.index("--samples") + 1] = "2" if generated else "3"
            command[command.index("--warmups") + 1] = "1" if generated else "2"
            full = ["taskset", "-c", CORE] + command
            completed = subprocess.run(full, capture_output=True, text=True, timeout=590)
            if completed.returncode != 0:
                raise SystemExit(f"{full} failed: {completed.stderr[-2000:]}")
            with open(f"{out_dir}/{name}-{arm}.json", "w") as handle:
                handle.write(completed.stdout)
            report = json.loads(completed.stdout)
            regions = report["allocation"]
            values = {field: sorted({region[field] for region in regions}) for field in FIELDS}
            summary.setdefault(name, {})[arm] = {
                "per_owner": {field: values[field][0] for field in FIELDS},
                "owners_identical": all(len(value) == 1 for value in values.values()),
                "output_sha256": report["output_sha256"],
            }
    json.dump({"schema": "0749-allocation-r2-v1", "summary": summary}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
