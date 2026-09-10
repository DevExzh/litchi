#!/usr/bin/env python3
"""Compare supplementary whole-child counters; no timed-operation attribution."""
from pathlib import Path
import csv, json
HERE = Path(__file__).resolve().parent
rows = []
for corpus in ["few-large", "many-small"]:
    data = {}
    for phase in ["before", "after"]:
        path = HERE / "profiles" / phase / f"{corpus}-owned-batch-w4.perf.csv"
        data[phase] = {r[2]: float(r[0]) for r in csv.reader(path.open())
                       if len(r) > 3 and r[0] and not r[0].startswith("#")}
    metrics = {key: {"before": value, "after": data["after"][key],
                     "delta_pct": (data["after"][key] / value - 1) * 100}
               for key, value in data["before"].items()}
    rows.append({"corpus": corpus, "source": "owned", "workers": 4, "metrics": metrics})
(HERE / "profile-comparison.json").write_text(json.dumps({
    "scope": "one whole-child profile per phase/corpus, includes setup and verification, not operation attribution",
    "rows": rows}, indent=2) + "\n")
