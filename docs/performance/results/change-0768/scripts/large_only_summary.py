#!/usr/bin/env python3
"""Summarize the large-only runs: per-pair p50 and page faults, B/A ratios.

Usage: large_only_summary.py DIR PREFIX_JSON PREFIX_PERF ROUNDS
e.g.   large_only_summary.py runs-tunables harness-large perf-large 4
"""
import csv
import json
import statistics
import sys

directory, json_prefix, perf_prefix, rounds = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4])
ratios = []
for r in range(rounds):
    values = {}
    for leg in "AB":
        report = json.load(open(f"{directory}/{json_prefix}-r{r}-{leg}.json"))
        rows = {
            row[2]: int(row[0])
            for row in csv.reader(l for l in open(f"{directory}/{perf_prefix}-r{r}-{leg}.csv") if l.strip() and not l.startswith("#"))
        }
        values[leg] = (report["results"][0]["elapsed_ns"]["p50"], rows["page-faults"], rows["instructions"])
    ratio = values["B"][0] / values["A"][0]
    ratios.append(ratio)
    print(f"r{r}: A p50 {values['A'][0]} ns, {values['A'][1]} page faults | B p50 {values['B'][0]} ns, {values['B'][1]} page faults | B/A {ratio:.3f}")
print(f"median B/A {statistics.median(ratios):.3f}")
