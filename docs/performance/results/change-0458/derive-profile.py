#!/usr/bin/env python3
"""Classify sampled stack periods by the harness's non-inlined phase markers."""
import collections
import csv
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
HEADER = re.compile(r"^\S.*:\s+(\d+)\s+cycles:u:\s*$")
MARKER = re.compile(r"odp_append_attribution::phase_(snapshot_open|transaction|add|commit|publication)(?:\W|$)")


def derive():
    path = ROOT / "profiling/samples/perf-script.stdout"
    groups = collections.defaultdict(lambda: {"samples": 0, "period": 0, "leaves": collections.Counter()})
    frames = []
    period = None
    malformed = 0

    def flush():
        nonlocal frames, period, malformed
        if period is None:
            return
        if not frames:
            malformed += 1
        else:
            markers = {match.group(1) for frame in frames for match in MARKER.finditer(frame)}
            group = next(iter(markers)) if len(markers) == 1 else ("unattributed" if not markers else "ambiguous")
            value = groups[group]
            value["samples"] += 1
            value["period"] += period
            leaf = re.sub(r"^[0-9a-f]+\s+", "", frames[0]).split(" (")[0]
            value["leaves"][leaf] += period
        frames = []
        period = None

    for line in path.read_text(errors="replace").splitlines():
        match = HEADER.match(line)
        if match:
            flush()
            period = int(match.group(1))
        elif not line.strip():
            flush()
        elif line[:1].isspace() and period is not None:
            frames.append(line.strip())
    flush()
    total = sum(row["period"] for row in groups.values())
    assert total > 0, "no sampled cycles:u stacks parsed"
    samples = sum(row["samples"] for row in groups.values())
    result = {}
    for name, value in sorted(groups.items()):
        result[name] = {"samples": value["samples"], "period": value["period"], "share_total_period_pct": value["period"] / total * 100,
                        "top_leaf_symbols": [{"symbol": symbol, "period": count, "share_phase_period_pct": count / value["period"] * 100}
                                             for symbol, count in value["leaves"].most_common(15)]}
    counters = {}
    with (ROOT / "profiling/counters/counters.csv").open() as stream:
        for row in csv.reader(stream):
            if len(row) >= 5 and row[2] in ["cycles", "instructions", "branches", "branch-misses", "cache-misses", "page-faults", "context-switches"]:
                counters[row[2]] = {"value": row[0], "unit": row[1], "event_runtime": row[3], "running_pct": row[4]}
    assert len(counters) == 7, "missing counter output rows"
    ipc = None
    if counters["cycles"]["value"].isdigit() and counters["instructions"]["value"].isdigit():
        ipc = int(counters["instructions"]["value"]) / int(counters["cycles"]["value"])
    return {"schema": "litchi-0458-profile-derived-v1", "change": 458, "parsed_samples": samples,
            "malformed_samples": malformed, "total_period": total, "groups": result,
            "counter_scope": "separate whole-process run including setup, warmups, checks and reporting; values retain perf's scaling/runtime fields",
            "counters": counters, "whole_process_ipc": ipc,
            "sampling_scope": "whole-process cycles:u sample periods; phase markers can include warmup calls; unresolved/truncated stacks remain unattributed; this is not hard-counter phase attribution"}


if __name__ == "__main__":
    print(json.dumps(derive(), indent=2, sort_keys=True))
