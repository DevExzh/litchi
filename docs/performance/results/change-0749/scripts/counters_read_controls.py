#!/usr/bin/env python3
"""Change 0749 per-owner counters for the harness CFB read controls.

The timing matrix showed `cfb_open/few-large` at +4.8% although no code on
the reader's open path changed. This lane measures the same per-owner
instruction and cycle counts as `counters.py` (two sample counts, three
argv[0] layouts) for the harness read controls, so a work change can be told
apart from a code-layout effect.

Usage: counters_read_controls.py OUT_DIR
"""

import json
import os
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from counters import EVENTS, LAYOUTS, measure  # noqa: E402
from run import ARMS, binary  # noqa: E402

CONTROLS = [
    (case, shape)
    for case in ("cfb_open", "cfb_read_one")
    for shape in ("few-large", "many-small")
]
SIZES = (40, 240)


def command(case, shape, arm, layout, samples, out):
    return [
        binary(arm, "litchi-perf-baseline", layout),
        "--case", case,
        "--shape", shape,
        "--samples", str(samples),
        "--warmup", "5",
        "--json", out,
    ]


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    rows = []
    for layout in LAYOUTS:
        for case, shape in CONTROLS:
            label = f"{case}/{shape}"
            for arm in ARMS:
                measured = {}
                for samples in SIZES:
                    stem = f"{out_dir}/{case}__{shape}-{arm}-l{layout}-n{samples}"
                    measured[samples] = measure(stem, command(case, shape, arm, layout, samples, f"{stem}.json"))
                per_owner = {
                    event: (measured[SIZES[1]][event] - measured[SIZES[0]][event]) / (SIZES[1] - SIZES[0])
                    for event in measured[SIZES[1]]
                }
                rows.append({"case": label, "arm": arm, "argv0_extra_bytes": 8 * layout,
                             "sizes": SIZES, "per_owner": per_owner})
    summary = {}
    for row in rows:
        entry = summary.setdefault(row["case"], {})
        for event, value in row["per_owner"].items():
            entry.setdefault(event, {}).setdefault(row["arm"], []).append(round(value, 1))
    for events in summary.values():
        for arms in events.values():
            base = statistics.median(arms["base"])
            cand = statistics.median(arms["cand"])
            arms["median_change_pct"] = round(100.0 * (cand / base - 1.0), 3) if base else None
    json.dump({"events": EVENTS, "layouts_argv0_extra_bytes": [8 * x for x in LAYOUTS],
               "rows": rows, "summary": summary}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
