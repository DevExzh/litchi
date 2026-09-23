#!/usr/bin/env python3
"""Change 0750 follow-up: tabulate scripts/run_dos.sh's results.jsonl: per case
and leg, the median seconds of its audits and the verdict."""
import json, statistics, sys

rows = [json.loads(line) for line in open(sys.argv[1])]
cases = list(dict.fromkeys(row["case"] for row in rows))
by = {(row["case"], row["leg"]): row for row in rows}
print("| case | bytes | base s | first candidate s | fix s | fix / base | verdicts |")
print("| --- | ---: | ---: | ---: | ---: | ---: | --- |")
for case in cases:
    cells = []
    verdicts = set()
    for leg in ("before", "prefix", "fixed"):
        row = by[(case, leg)]
        seconds = statistics.median(row["seconds"])
        runs = len(row["seconds"])
        cells.append((seconds, runs))
        verdicts.add(row["verdict"].split("(")[0] if row["verdict"].startswith("Ok(") else row["verdict"])
    (base, _), (first, first_runs), (fixed, _) = cells
    note = "" if first_runs > 1 else " (1 run)"
    verdict = " / ".join(sorted(verdicts)) + (" on all legs" if len(verdicts) == 1 else "")
    print(f"| `{case}` | {by[(case, 'fixed')]['bytes']:,} | {base:.4f} | {first:.3f}{note} | {fixed:.4f} | {fixed / base:.2f} | {verdict} |")
