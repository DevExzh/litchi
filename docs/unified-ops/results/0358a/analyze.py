#!/usr/bin/env python3
"""Summarize the record 0358a A-B-B-A runs: per probe and scenario, the median
of the per-run medians on each leg, after/before, the peak heap growth, and
whether every leg wrote byte-identical outputs (the dumped first-round
outputs, compared by SHA-256)."""
import hashlib, json, pathlib, statistics, sys

base = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else pathlib.Path(__file__).parent)
summary = {}
for probe in ("sb", "w6"):
    legs = {"before": {}, "after": {}}
    heap = {"before": {}, "after": {}}
    digest = {"before": {}, "after": {}}
    for path in sorted(base.glob(f"{probe}-*-*.json")):
        leg = path.name.split("-")[1]
        for row in json.loads(path.read_text()):
            legs[leg].setdefault(row["scenario"], []).append(row["median_us"])
            if "peak_heap_bytes" in row:
                heap[leg].setdefault(row["scenario"], []).append(row["peak_heap_bytes"])
                digest[leg].setdefault(row["scenario"], set()).add(
                    (row["output_bytes"], row["output_siphash"]))
    rows = []
    for scenario, after in legs["after"].items():
        before = legs["before"].get(scenario)
        if not before:
            continue
        b, a = statistics.median(before), statistics.median(after)
        entry = {"scenario": scenario, "before_us": round(b, 1), "after_us": round(a, 1),
                 "ratio": round(a / b, 3), "before_runs": before, "after_runs": after}
        if scenario in heap["after"]:
            hb, ha = max(heap["before"][scenario]), max(heap["after"][scenario])
            entry.update(before_peak_heap_bytes=hb, after_peak_heap_bytes=ha,
                         peak_heap_ratio=round(ha / hb, 3))
            outputs = digest["before"][scenario] | digest["after"][scenario]
            entry["outputs_identical_siphash"] = len(outputs) == 1
        rows.append(entry)
    summary[probe] = rows
dumps = {}
for directory in sorted(base.glob("sb-*-*")):
    if directory.is_dir():
        for file in directory.iterdir():
            dumps.setdefault(file.name, set()).add(hashlib.sha256(file.read_bytes()).hexdigest())
summary["sb_outputs_sha256_identical"] = {name: len(hashes) == 1 for name, hashes in sorted(dumps.items())}
summary["sb_outputs_sha256"] = {name: sorted(hashes) for name, hashes in sorted(dumps.items())}
print(json.dumps(summary, indent=1))
