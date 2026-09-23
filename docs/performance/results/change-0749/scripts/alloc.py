#!/usr/bin/env python3
"""Change 0749 counting-allocator lane.

One pinned `probe_alloc` process per probe case and arm, 3 warmups and 5
measured owners. Each owner's allocation region (bytes requested and
returned, successful calls, peak live bytes relative to entry, retained bytes)
is kept; the summary reports the per-owner values, which are identical across
owners of one process for these deterministic workloads, and flags any case
where they are not.

Usage: alloc.py OUT_DIR > allocation.json
"""

import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from run import ARMS, CASES, CORE, binary  # noqa: E402

FIELDS = ("allocated_bytes", "allocation_calls", "peak_live_bytes", "retained_bytes")


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    summary = {}
    for name in CASES:
        if name.startswith("harness-"):
            continue
        for arm in ARMS:
            command = CASES[name](arm, 0, "/dev/null")
            command[0] = binary(arm, "probe_alloc", 0)
            command[command.index("--samples") + 1] = "5"
            command[command.index("--warmups") + 1] = "3"
            full = ["taskset", "-c", CORE] + command
            completed = subprocess.run(full, capture_output=True, text=True)
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
        base = summary[name]["base"]["per_owner"]
        cand = summary[name]["cand"]["per_owner"]
        summary[name]["change"] = {field: cand[field] - base[field] for field in FIELDS}
    json.dump({"schema": "0749-allocation-v1", "summary": summary}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
