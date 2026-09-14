#!/usr/bin/env python3
"""Folds the change-0576 captures into the numbers the record cites.

Reads only this directory. Pure standard library.
"""

import glob
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROUNDS = ("a1", "b1", "b2", "a2")
LEG_OF = {"a1": "before", "a2": "before", "b1": "after", "b2": "after"}
FIXTURES = ("flagship", "cv", "54016")


def percentile(values, fraction):
    ordered = sorted(values)
    if not ordered:
        return float("nan")
    position = (len(ordered) - 1) * fraction
    low = int(position)
    high = min(low + 1, len(ordered) - 1)
    weight = position - low
    return ordered[low] * (1 - weight) + ordered[high] * weight


SUMMARY_NAME = "latency-summary.json"


def locate(root, subdirectory, name):
    """Retained evidence keeps perf CSVs and callgrind text in subdirectories;
    a freshly captured scratch directory keeps them flat. Accept both."""
    nested = os.path.join(root, subdirectory, name)
    return nested if os.path.exists(nested) else os.path.join(root, name)


def load_latency(root):
    """Loads the folded summary if one is present, else the raw captures.

    The raw captures are 41 MB of per-sample arrays and are not retained; the
    fold keeps the statistics, the logical counters and the binary identity,
    which is everything the record cites.
    """
    folded = os.path.join(root, SUMMARY_NAME)
    if os.path.exists(folded):
        document = json.load(open(folded, encoding="utf-8"))
        return {
            (cell["round"], cell["fixture"], cell["mode"], cell["operation"]): {
                "p50": cell["p50"],
                "p90": cell["p90"],
                "p99": cell["p99"],
                "mean": cell["mean"],
                "n": cell["n"],
                "counters": tuple(cell["counters"]),
                "sha256": cell["sha256"],
            }
            for cell in document["cells"]
        }
    return load_raw_latency(root)


def load_raw_latency(root):
    cells = {}
    for round_name in ROUNDS:
        for fixture in FIXTURES:
            pattern = os.path.join(root, round_name, fixture, "*.json")
            for path in sorted(glob.glob(pattern)):
                stem = os.path.basename(path)[: -len(".json")]
                for candidate in ("one-cell", "list", "open"):
                    if stem.endswith("-" + candidate):
                        mode = stem[: -len(candidate) - 1]
                        operation = candidate
                        break
                else:
                    continue
                try:
                    report = json.load(open(path, encoding="utf-8"))
                except json.JSONDecodeError:
                    # A cell the driver could not produce leaves an empty file.
                    # `54016.xls` has fewer worksheets than the driver's fixed
                    # `--worksheet-index 1`, so its one-cell operation is absent.
                    continue
                samples = report["elapsed_samples_ns"]
                counters = {
                    (
                        record["metrics"]["read_calls"],
                        record["metrics"]["read_bytes"],
                        record["metrics"]["version_calls"],
                    )
                    for record in report["records"]
                }
                assert len(counters) == 1, f"{path} has varying logical counters"
                cells[(round_name, fixture, mode, operation)] = {
                    "p50": percentile(samples, 0.5),
                    "p90": percentile(samples, 0.9),
                    "p99": percentile(samples, 0.99),
                    "mean": sum(samples) / len(samples),
                    "n": len(samples),
                    "counters": next(iter(counters)),
                    "sha256": report["binary"]["sha256"],
                }
    return cells


def latency_tables(root):
    cells = load_latency(root)
    if not cells:
        print("no latency captures found")
        return
    keys = sorted(
        {
            (fixture, mode, operation)
            for (_, fixture, mode, operation) in cells
            if all(
                (round_name, fixture, mode, operation) in cells for round_name in ROUNDS
            )
        }
    )
    partial = sorted(
        {
            (fixture, mode, operation)
            for (_, fixture, mode, operation) in cells
            if not all(
                (round_name, fixture, mode, operation) in cells for round_name in ROUNDS
            )
        }
    )
    if partial:
        print(f"cells missing from at least one round (excluded): {partial}")

    print("== logical counters, both legs ==")
    mismatches = 0
    for fixture, mode, operation in keys:
        counters = {
            LEG_OF[round_name]: cells[(round_name, fixture, mode, operation)]["counters"]
            for round_name in ROUNDS
            if (round_name, fixture, mode, operation) in cells
        }
        same = len(set(counters.values())) == 1
        mismatches += 0 if same else 1
        reads, byte_count, versions = counters["before"]
        print(
            f"{fixture:9s} {mode:12s} {operation:9s} reads={reads:5d} bytes={byte_count:9d} "
            f"version={versions:4d} identical={'yes' if same else 'NO'}"
        )
    print(f"counter mismatches: {mismatches}")

    print()
    print("== paired medians, two directions ==")
    print(
        f"{'fixture':9s} {'mode':12s} {'op':9s} {'before p50':>12s} {'after p50':>12s} "
        f"{'dir1':>9s} {'dir2':>9s} {'A/A':>8s} {'B/B':>8s}"
    )
    for fixture, mode, operation in keys:
        cell = lambda r: cells[(r, fixture, mode, operation)]["p50"]
        a1, b1, b2, a2 = (cell(r) for r in ROUNDS)
        direction1 = (b1 - a1) / a1 * 100
        direction2 = (b2 - a2) / a2 * 100
        noise_a = (a2 - a1) / a1 * 100
        noise_b = (b2 - b1) / b1 * 100
        before = (a1 + a2) / 2
        after = (b1 + b2) / 2
        print(
            f"{fixture:9s} {mode:12s} {operation:9s} {before:12,.0f} {after:12,.0f} "
            f"{direction1:+8.2f}% {direction2:+8.2f}% {noise_a:+7.2f}% {noise_b:+7.2f}%"
        )

    print()
    print("== p90 / p99, paired ==")
    for fixture, mode, operation in keys:
        row = []
        for statistic in ("p90", "p99"):
            values = [
                cells[(r, fixture, mode, operation)][statistic] for r in ROUNDS
            ]
            before = (values[0] + values[3]) / 2
            after = (values[1] + values[2]) / 2
            row.append(f"{statistic} {before:10,.0f} -> {after:10,.0f} ({(after - before) / before * 100:+6.2f}%)")
        print(f"{fixture:9s} {mode:12s} {operation:9s} " + "  ".join(row))


PERF_EVENTS = ("instructions", "branches", "branch-misses", "task-clock")


def parse_perf(path):
    values = {}
    for line in open(path, encoding="utf-8"):
        if line.startswith("#") or not line.strip():
            continue
        fields = line.strip().split(",")
        if len(fields) < 3:
            continue
        try:
            values[fields[2]] = float(fields[0])
        except ValueError:
            continue
    return values


def perf_table(root):
    print()
    print("== hardware counters per open, isolated by differencing 1100 and 100 samples ==")
    print(f"{'fixture':9s} {'leg':7s} " + " ".join(f"{event:>16s}" for event in PERF_EVENTS))
    per_fixture = {}
    for fixture in FIXTURES:
        for leg in ("before", "after"):
            small = locate(root, "counters", f"perf-{leg}-{fixture}-s100.csv")
            large = locate(root, "counters", f"perf-{leg}-{fixture}-s1100.csv")
            if not (os.path.exists(small) and os.path.exists(large)):
                continue
            low, high = parse_perf(small), parse_perf(large)
            row = {
                event: (high.get(event, 0) - low.get(event, 0)) / 1000
                for event in PERF_EVENTS
            }
            per_fixture.setdefault(fixture, {})[leg] = row
            print(
                f"{fixture:9s} {leg:7s} "
                + " ".join(f"{row[event]:16,.0f}" for event in PERF_EVENTS)
            )
    print()
    for fixture, legs in per_fixture.items():
        if "before" in legs and "after" in legs:
            for event in PERF_EVENTS:
                before, after = legs["before"][event], legs["after"][event]
                if before:
                    print(
                        f"{fixture:9s} {event:14s} {before:14,.0f} -> {after:14,.0f} "
                        f"({(after - before) / before * 100:+6.2f}%)"
                    )


ANN_LINE = re.compile(r"^\s*([\d,]+)(?:\s*\([^)]*\))?\s+(\S.*?)\s*$")
# `callgrind_annotate` suffixes every symbol with the object file it came from.
# The two legs are different files, so the suffix has to go before the two
# profiles can be compared symbol by symbol.
OBJECT_SUFFIX = re.compile(r"\s*\[[^\]]*\]\s*$")
SST_MARKS = ("from_utf16", "SstCursor", "shared_string", "SharedString", "walk_one")
CALLGRIND_CELLS = {"flagship": 200, "cv": 100, "54016": 50}


def annotate(path):
    totals = {}
    for line in open(path, encoding="utf-8", errors="replace"):
        match = ANN_LINE.match(line)
        if not match:
            continue
        name = OBJECT_SUFFIX.sub("", match.group(2))
        if name.startswith("Ir") or "PROGRAM TOTALS" in name or name.startswith("--"):
            continue
        try:
            count = int(match.group(1).replace(",", ""))
        except ValueError:
            continue
        totals[name] = totals.get(name, 0) + count
    return totals


def callgrind_table(root):
    print()
    print("== callgrind self cost per open ==")
    for fixture, opens in CALLGRIND_CELLS.items():
        legs = {}
        for leg in ("before", "after"):
            small = locate(root, "callgrind", f"ann-{leg}-{fixture}-ssmall.txt")
            large = locate(root, "callgrind", f"ann-{leg}-{fixture}-slarge.txt")
            if not (os.path.exists(small) and os.path.exists(large)):
                continue
            low, high = annotate(small), annotate(large)
            delta = {}
            for name in set(low) | set(high):
                value = (high.get(name, 0) - low.get(name, 0)) / opens
                if abs(value) >= 1:
                    delta[name] = value
            legs[leg] = delta
        if len(legs) != 2:
            continue
        for leg, delta in legs.items():
            total = sum(delta.values())
            sst = sum(v for k, v in delta.items() if any(m in k for m in SST_MARKS))
            print(
                f"{fixture:9s} {leg:7s} total Ir/open={total:12,.0f} "
                f"shared-string self={sst:12,.0f} ({sst / total * 100:5.2f}%)"
            )
        before_total = sum(legs["before"].values())
        after_total = sum(legs["after"].values())
        print(
            f"{fixture:9s} delta   total Ir/open {before_total:,.0f} -> {after_total:,.0f} "
            f"({(after_total - before_total) / before_total * 100:+6.2f}%)"
        )
        names = sorted(
            set(legs["before"]) | set(legs["after"]),
            key=lambda n: (-abs(legs["after"].get(n, 0) - legs["before"].get(n, 0)), n),
        )
        for name in names[:8]:
            before = legs["before"].get(name, 0)
            after = legs["after"].get(name, 0)
            print(f"    {name[:78]:78s} {before:12,.0f} -> {after:12,.0f}")

    print()
    print("== callgrind inclusive cost of the SST scan per open ==")
    for fixture, opens in CALLGRIND_CELLS.items():
        for leg in ("before", "after"):
            small = locate(root, "callgrind", f"incl-{leg}-{fixture}-ssmall.txt")
            large = locate(root, "callgrind", f"incl-{leg}-{fixture}-slarge.txt")
            if not (os.path.exists(small) and os.path.exists(large)):
                continue
            low, high = annotate(small), annotate(large)
            scan = 0
            for name in set(low) | set(high):
                if "scan_shared_string_records" in name:
                    scan = max(scan, (high.get(name, 0) - low.get(name, 0)) / opens)
            print(f"{fixture:9s} {leg:7s} scan_shared_string_records inclusive Ir/open={scan:12,.0f}")


def fold(root, destination, environment):
    cells = load_raw_latency(root)
    document = {
        "environment": environment,
        "cells": [
            {
                "round": round_name,
                "fixture": fixture,
                "mode": mode,
                "operation": operation,
                **{key: value for key, value in body.items() if key != "counters"},
                "counters": list(body["counters"]),
            }
            for (round_name, fixture, mode, operation), body in sorted(cells.items())
        ],
    }
    json.dump(document, open(destination, "w", encoding="utf-8"), indent=1, sort_keys=True)
    print(f"folded {len(document['cells'])} cells into {destination}")


def main():
    arguments = sys.argv[1:]
    if arguments[:1] == ["--fold"]:
        fold(arguments[1], arguments[2], json.load(open(arguments[3], encoding="utf-8")))
        return
    root = arguments[0] if arguments else HERE
    latency_tables(root)
    perf_table(root)
    callgrind_table(root)


if __name__ == "__main__":
    main()
