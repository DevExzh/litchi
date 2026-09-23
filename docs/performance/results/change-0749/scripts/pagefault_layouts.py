#!/usr/bin/env python3
"""Change 0749: page faults and cycles per owner across ten heap layouts.

`doc_semantic_one_edit_save/large` showed ~500 more page faults per owner
for the candidate at the three counter layouts, while a direct run showed the
opposite. This lane measures both arms at argv[0] extra lengths 0..72 bytes
(10 layouts) with the per-owner differencing of counters.py (20 and 220
samples), to tell a systematic effect from an allocator-state effect.

With `GLIBC_TUNABLES` set in the environment (for example
`glibc.malloc.trim_threshold=268435456:glibc.malloc.mmap_threshold=268435456`)
the same lane tests whether fixing glibc's heap-trim and mmap thresholds
removes the difference, which would identify the allocator policy rather
than the measured work as its cause. `CASES` may be narrowed with a
`PF_CASES` comma list of `case/shape` labels.

Usage: pagefault_layouts.py OUT_DIR > pagefaults.json
"""

import json
import os
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from counters import harness_command, measure  # noqa: E402
from run import ARMS  # noqa: E402

CASES = [("doc_semantic_one_edit_save", "large"), ("doc_semantic_one_edit_save", "tiny")]
SIZES = (20, 220)


def main():
    out_dir = sys.argv[1]
    selected = os.environ.get("PF_CASES")
    cases = [c for c in CASES if not selected or f"{c[0]}/{c[1]}" in selected.split(",")]
    os.makedirs(out_dir, exist_ok=True)
    rows = []
    for layout in range(10):
        for case, shape in cases:
            for arm in ARMS:
                measured = {}
                for samples in SIZES:
                    stem = f"{out_dir}/{case}__{shape}-{arm}-l{layout}-n{samples}"
                    measured[samples] = measure(stem, harness_command(case, shape, arm, layout, samples, f"{stem}.json"))
                per_owner = {event: (measured[SIZES[1]][event] - measured[SIZES[0]][event]) / (SIZES[1] - SIZES[0])
                             for event in measured[SIZES[1]]}
                rows.append({"case": f"{case}/{shape}", "arm": arm, "argv0_extra_bytes": 8 * layout, "per_owner": per_owner})
    summary = {}
    for row in rows:
        entry = summary.setdefault(row["case"], {}).setdefault(row["arm"], {})
        for event, value in row["per_owner"].items():
            entry.setdefault(event, []).append(round(value, 1))
    json.dump({"sizes": SIZES, "glibc_tunables": os.environ.get("GLIBC_TUNABLES"),
               "rows": rows, "summary": summary}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
