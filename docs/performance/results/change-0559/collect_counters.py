#!/usr/bin/env python3
"""Collect the retained Callgrind branch-simulation summaries into one table.

Each summary line is the deterministic whole-child total for one case and one
build stage. Corpus generation is inside the child, so these are whole-child
counters, not owner-scoped attribution.
"""

from __future__ import annotations

import argparse
import json
import pathlib
import re

EVENTS = ("Ir", "Bc", "Bcm", "Bi", "Bim")


def summary(path: pathlib.Path) -> dict[str, int]:
    for line in path.read_text(errors="replace").splitlines():
        if line.startswith("summary:"):
            values = [int(value) for value in re.findall(r"\d+", line)]
            return dict(zip(EVENTS, values, strict=False))
    raise SystemExit(f"{path}: no summary line")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile-dir", required=True)
    parser.add_argument("--output", default="docs/performance/results/change-0559/counters.json")
    args = parser.parse_args()
    root = pathlib.Path(args.profile_dir)

    rows = []
    for baseline in sorted(root.glob("*-A.callgrind")):
        candidate = baseline.with_name(baseline.name[: -len("-A.callgrind")] + "-C.callgrind")
        if not candidate.exists():
            continue
        name = baseline.name[: -len("-A.callgrind")]
        before, after = summary(baseline), summary(candidate)
        rows.append(
            {
                "profile": name,
                "baseline": before,
                "candidate": after,
                "change_percent": {
                    event: (after[event] - before[event]) / before[event] * 100.0
                    for event in EVENTS
                    if before.get(event)
                },
            }
        )

    pathlib.Path(args.output).write_text(json.dumps({"schema_version": 1, "events": list(EVENTS), "rows": rows}, indent=2) + "\n")
    header = f"{'profile':34s}" + "".join(f"{event:>12s}" for event in EVENTS)
    print(header)
    for row in rows:
        print(f"{row['profile']:34s}" + "".join(f"{row['change_percent'].get(event, 0.0):+11.2f}%" for event in EVENTS))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
