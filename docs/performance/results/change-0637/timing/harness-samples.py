"""Extract per-sample nanoseconds for one case and shape from a harness report."""
import json, sys
report, case, shape = sys.argv[1], sys.argv[2], sys.argv[3]
with open(report) as handle:
    doc = json.load(handle)
for row in doc["results"]:
    if row["case"] != case:
        continue
    if shape != "*" and row.get("corpus", {}).get("shape") != shape:
        continue
    for sample in row["elapsed_ns"]["samples"]:
        print(sample)
