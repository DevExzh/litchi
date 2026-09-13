#!/usr/bin/env python3
"""Summarize a frozen complex-function capture without extracting scratch files."""
import json
import re
import statistics
import sys
import tarfile

with tarfile.open(sys.argv[1]) as archive:
    receipt = json.load(archive.extractfile("receipt.json"))
    assert receipt["binary_unchanged"]
    assert len(receipt["records"]) == 225
    assert all(record["status"] == 0 for record in receipt["records"])
    rows = []
    for case in receipt["cases"]:
        records = [r for r in receipt["records"] if r["case"] == case]
        results = [r["result"] for r in records]
        assert len(results) == 3
        rss = []
        for record in records:
            name = f"{record['round']:02}-{case}.time"
            raw = archive.extractfile(name).read().decode()
            rss.append(int(re.search(r"Maximum resident set size \(kbytes\): (\d+)", raw)[1]))
        rows.append({
            "case": case,
            "repeat": results[0]["repeat"],
            "median_ns_per_repeat": statistics.median(r["elapsed_ns_per_repeat"] for r in results),
            "p95_ns_per_repeat": statistics.median(r["elapsed_ns_p95"] / r["repeat"] for r in results),
            "rss_kib": rss,
            "allocator_calls_per_repeat": results[0]["allocator_calls"] / results[0]["repeat"],
            "peak_live_delta": results[0]["peak_live_delta"],
            "retained_bytes": results[0]["memory_retained"],
            "work_per_repeat": results[0]["work_per_repeat"],
            "expected_failure": results[0]["expected_failure"],
        })
print(json.dumps({"runs": 225, "cases": rows}, indent=2))
