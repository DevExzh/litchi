#!/usr/bin/env python3
"""Recomputes every table change 0605 cites, from this directory alone.

    python3 -B docs/performance/results/change-0605/analyze.py \
      docs/performance/results/change-0605

Standard library only. Folded inputs, all retained here:

  counters-summary.json   logical counters per (fixture, mode, operation), leg
  walk-summary.json       the all-cells scan against the per-cell reading
  latency-summary.json    the A1 B1 B2 A2 wall-clock rounds
  callgrind/ann-<leg>-<fixture>-<label>-s{small,large}.txt
                          callgrind_annotate self cost, the isolation pair
                          change 0574 defined
  perf/perf-<leg>-<fixture>-<label>-s<n>.csv
                          perf stat -x, hardware counters, isolated the same way

`--fold <capture-dir>` rebuilds the three summaries from a raw capture, which is
how they were produced; the raw per-sample JSON is not retained because the
folded form carries every number the record cites and the raw form is 9 MB.

The callgrind and perf figures are differences between a large-sample and a
small-sample child divided by the extra operations, so everything that runs once
per child -- the harness's SHA-256 of the input, its oracle projection, which for
all-cells includes a whole differential walk -- cancels.
"""

import json
import os
import re
import statistics
import sys

STEMS = ["flagship", "cv", "54016"]
MODES = ["owned-readat", "file-source"]
BOTH_LEG_OPS = ["open", "list", "one-cell"]
AFTER_ONLY_OPS = ["second-cell", "full-text", "all-cells", "all-cells-per-cell"]
COUNTERS = ["read_calls", "read_bytes", "version_calls", "len_calls", "seek_calls"]

# (fixture, label): (small, large) sample counts, matching capture_callgrind.sh
CALLGRIND_SAMPLES = {
    ("flagship", "open"): (20, 220),
    ("flagship", "one-cell"): (20, 220),
    ("cv", "open"): (20, 120),
    ("cv", "one-cell"): (20, 120),
    ("54016", "open"): (10, 60),
    ("54016", "one-cell"): (5, 25),
    ("flagship", "all-cells-scan"): (5, 25),
    ("flagship", "all-cells-per-cell"): (5, 25),
    ("cv", "all-cells-scan"): (2, 6),
    ("cv", "all-cells-per-cell"): (2, 6),
    ("54016", "all-cells-scan"): (2, 6),
    ("54016", "all-cells-per-cell"): (2, 6),
}
# (fixture, label): (small, large), matching capture_perf.sh
PERF_SAMPLES = {
    (stem, label): (100, 1100) for stem in STEMS for label in ("open", "one-cell")
}
PERF_SAMPLES.update(
    {
        ("flagship", "all-cells-scan"): (20, 220),
        ("flagship", "all-cells-per-cell"): (20, 220),
        ("cv", "all-cells-scan"): (5, 55),
        ("cv", "all-cells-per-cell"): (5, 55),
        ("54016", "all-cells-scan"): (5, 55),
        ("54016", "all-cells-per-cell"): (5, 55),
    }
)
LINE = re.compile(r"^\s*([\d,]+) \([\s\d.]+%\)\s+(.*?)(?: \[[^\]]*\])?$")


# --------------------------------------------------------------------------
# callgrind


def annotate(path):
    """Self Ir per symbol, plus PROGRAM TOTALS, from one callgrind_annotate."""
    totals, symbols, body = None, {}, False
    with open(path, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            if "PROGRAM TOTALS" in line:
                totals = int(line.split("(")[0].strip().replace(",", ""))
                continue
            if line.startswith("Ir") and "file:function" in line:
                body = True
                continue
            if not body:
                continue
            match = LINE.match(line.rstrip("\n"))
            if not match:
                continue
            name = match.group(2).strip()
            if name in ("PROGRAM TOTALS", "file:function"):
                continue
            symbols[name] = symbols.get(name, 0) + int(match.group(1).replace(",", ""))
    return totals, symbols


def isolate(root, leg, stem, label):
    """One operation's Ir: (large - small) / (large samples - small samples)."""
    small, large = CALLGRIND_SAMPLES[(stem, label)]
    paths = [
        os.path.join(root, "callgrind", f"ann-{leg}-{stem}-{label}-s{size}.txt")
        for size in ("small", "large")
    ]
    if not all(os.path.exists(path) for path in paths):
        return None, {}
    (small_total, small_symbols), (large_total, large_symbols) = (
        annotate(paths[0]),
        annotate(paths[1]),
    )
    span = large - small
    per = {
        name: (large_symbols.get(name, 0) - small_symbols.get(name, 0)) / span
        for name in set(small_symbols) | set(large_symbols)
    }
    return (large_total - small_total) / span, per


# --------------------------------------------------------------------------
# perf


def perf_read(path):
    values = {}
    with open(path, encoding="utf-8") as handle:
        for line in handle:
            if line.startswith("#") or not line.strip():
                continue
            parts = line.split(",")
            if len(parts) < 3:
                continue
            try:
                value = float(parts[0])
            except ValueError:
                continue
            values[parts[2].strip()] = value
    return values


def perf_isolate(root, leg, stem, label):
    small, large = PERF_SAMPLES[(stem, label)]
    paths = [
        os.path.join(root, "perf", f"perf-{leg}-{stem}-{label}-s{n}.csv")
        for n in (small, large)
    ]
    if not all(os.path.exists(path) for path in paths):
        return None
    a, b = perf_read(paths[0]), perf_read(paths[1])
    span = large - small
    return {
        key: (b.get(key, 0) - a.get(key, 0)) / span
        for key in ("cycles", "instructions", "branches", "branch-misses")
    }


# --------------------------------------------------------------------------
# folding a raw capture


def statistics_of(samples):
    ordered = sorted(samples)
    n = len(ordered)

    def percentile(fraction):
        if n == 1:
            return ordered[0]
        index = min(n - 1, max(0, int(round(fraction * (n - 1)))))
        return ordered[index]

    return {
        "n": n,
        "p50": statistics.median(ordered),
        "mean": statistics.fmean(ordered),
        "p95": percentile(0.95),
        "p99": percentile(0.99),
    }


def fold(capture, root):
    cells = []
    counters_dir = os.path.join(capture, "counters")
    for stem in STEMS:
        for mode in MODES:
            for op in BOTH_LEG_OPS + AFTER_ONLY_OPS:
                cell = {"fixture": stem, "mode": mode, "operation": op, "legs": {}}
                for leg in ("before", "after"):
                    path = os.path.join(counters_dir, f"{leg}-{stem}-{mode}-{op}.json")
                    if not os.path.exists(path):
                        continue
                    with open(path, encoding="utf-8") as handle:
                        data = json.load(handle)
                    first = data["records"][0]["metrics"]
                    for record in data["records"]:
                        for key in COUNTERS:
                            assert record["metrics"][key] == first[key], (
                                f"{leg}/{stem}/{mode}/{op}: {key} varies between samples"
                            )
                    cell["legs"][leg] = {
                        **{key: first[key] for key in COUNTERS},
                        "observation": data["records"][0]["observation"],
                        "binary_sha256": data["binary"]["sha256"],
                    }
                if cell["legs"]:
                    cells.append(cell)
    write(os.path.join(root, "counters-summary.json"), {"cells": cells})

    walks = []
    walk_dir = os.path.join(capture, "walk")
    if os.path.isdir(walk_dir):
        for name in sorted(os.listdir(walk_dir)):
            if not name.endswith(".json"):
                continue
            with open(os.path.join(walk_dir, name), encoding="utf-8") as handle:
                data = json.load(handle)
            metrics = data["records"][0]["metrics"]
            walks.append(
                {
                    "case": name[len("walk-") : -len(".json")],
                    "worksheet_index": data["worksheet_index"],
                    "strategy": data["all_cells_strategy"],
                    "per_cell_limit": data["per_cell_limit"],
                    "cells_reported": data["records"][0]["observation"]["cells_reported"],
                    "outcome": data["records"][0]["observation"]["outcome"],
                    **{key: metrics[key] for key in COUNTERS},
                    "elapsed": statistics_of(data["elapsed_samples_ns"]),
                }
            )
    write(os.path.join(root, "walk-summary.json"), {"cases": walks})

    rounds = {}
    latency_dir = os.path.join(capture, "latency")
    for round_name in ("a1", "b1", "b2", "a2"):
        for stem in STEMS:
            directory = os.path.join(latency_dir, round_name, stem)
            if not os.path.isdir(directory):
                continue
            for name in sorted(os.listdir(directory)):
                if not name.endswith(".json"):
                    continue
                with open(os.path.join(directory, name), encoding="utf-8") as handle:
                    data = json.load(handle)
                key = f"{stem}/{name[:-len('.json')]}"
                rounds.setdefault(key, {})[round_name] = {
                    **statistics_of(data["elapsed_samples_ns"]),
                    "binary_sha256": data["binary"]["sha256"],
                }
    write(os.path.join(root, "latency-summary.json"), {"cells": rounds})


def write(path, payload):
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=1, sort_keys=True)
        handle.write("\n")


# --------------------------------------------------------------------------
# tables


def load(root, name):
    path = os.path.join(root, name)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def table_counters(root):
    data = load(root, "counters-summary.json")
    print("== Logical counters, before against after ==")
    print("Every existing operation must be identical in both legs.\n")
    header = f"{'fixture':10} {'mode':13} {'operation':19} {'leg':7} {'reads':>7} {'bytes':>11} {'vers':>7} {'cells':>7}"
    print(header)
    moved = 0
    for cell in data["cells"]:
        for leg in ("before", "after"):
            row = cell["legs"].get(leg)
            if not row:
                continue
            cells_reported = row["observation"].get("cells_reported")
            print(
                f"{cell['fixture']:10} {cell['mode']:13} {cell['operation']:19} {leg:7} "
                f"{row['read_calls']:7} {row['read_bytes']:11} {row['version_calls']:7} "
                f"{str(cells_reported if cells_reported is not None else '-'):>7}"
            )
        legs = cell["legs"]
        if "before" in legs and "after" in legs:
            for key in COUNTERS:
                if legs["before"][key] != legs["after"][key]:
                    moved += 1
                    print(f"  !! {cell['operation']}: {key} moved")
            # The after leg's report carries two fields the before leg's
            # schema did not have; compare what both legs actually reported.
            def reported(leg):
                return {
                    key: value
                    for key, value in legs[leg]["observation"].items()
                    if value is not None
                }

            if reported("before") != reported("after"):
                moved += 1
                print(f"  !! {cell['operation']}: observation moved")
    print(f"\ncounters that moved between legs: {moved}\n")


def table_walk(root):
    data = load(root, "walk-summary.json")
    print("== The whole-sheet walk against the per-cell reading it replaces ==")
    print("Counters are per operation; p50 is over the retained samples.\n")
    print(
        f"{'case':34} {'sheet':>5} {'strategy':9} {'limit':>5} {'cells':>6} "
        f"{'reads':>7} {'bytes':>11} {'vers':>7} {'p50 ns':>14}"
    )
    for case in data["cases"]:
        print(
            f"{case['case']:34} {case['worksheet_index']:5} {case['strategy']:9} "
            f"{case['per_cell_limit']:5} {case['cells_reported']:6} {case['read_calls']:7} "
            f"{case['read_bytes']:11} {case['version_calls']:7} {case['elapsed']['p50']:14,.0f}"
        )
    print()
    print("Bytes per reported cell, and the per-cell leg extrapolated to the sheet:\n")
    by_case = {case["case"]: case for case in data["cases"]}
    for stem in STEMS:
        for mode in MODES:
            scan = by_case.get(f"{stem}-{mode}-scan")
            per = by_case.get(f"{stem}-{mode}-per-cell-256")
            if not scan or not per or not scan["cells_reported"]:
                continue
            scan_per_cell = scan["read_bytes"] / scan["cells_reported"]
            per_per_cell = per["read_bytes"] / per["cells_reported"]
            projected = per["elapsed"]["p50"] * scan["cells_reported"] / per["cells_reported"]
            print(
                f"{stem:10} {mode:13} walk {scan_per_cell:10,.1f} B/cell   "
                f"per-cell {per_per_cell:12,.1f} B/cell   "
                f"ratio {per_per_cell / scan_per_cell:9,.0f}x   "
                f"projected per-cell p50 {projected / 1e6:10,.1f} ms "
                f"against the walk's {scan['elapsed']['p50'] / 1e6:,.2f} ms"
            )
    print()


def table_callgrind(root):
    print("== Instructions per operation, callgrind isolation pairs ==\n")
    print(f"{'fixture':10} {'operation':12} {'before':>14} {'after':>14} {'delta':>9}")
    for stem in STEMS:
        for label in ("open", "one-cell"):
            before, _ = isolate(root, "before", stem, label)
            after, _ = isolate(root, "after", stem, label)
            if before is None or after is None:
                continue
            print(
                f"{stem:10} {label:12} {before:14,.0f} {after:14,.0f} "
                f"{100 * (after - before) / before:+8.3f}%"
            )
    print()
    print("all-cells, after leg only: the walk against 64 per-cell queries\n")
    print(f"{'fixture':10} {'leg':22} {'Ir per operation':>18} {'Ir per cell':>14}")
    walk = load(root, "walk-summary.json")
    cells_by_stem = {}
    for case in (walk or {}).get("cases", []):
        if case["strategy"] == "scan" and case["case"].endswith("owned-readat-scan"):
            cells_by_stem[case["case"].split("-")[0]] = case["cells_reported"]
    for stem in STEMS:
        for label in ("all-cells-scan", "all-cells-per-cell"):
            value, _ = isolate(root, "after", stem, label)
            if value is None:
                continue
            cells = cells_by_stem.get(stem) if label.endswith("scan") else 64
            per_cell = f"{value / cells:14,.0f}" if cells else " " * 14
            print(f"{stem:10} {label:22} {value:18,.0f} {per_cell}")
    print()


def table_perf(root):
    print("== Cycles and IPC, native perf stat, same isolation method ==")
    print("Single-shot counters: they carry the host's variance, unlike Ir.\n")
    print(
        f"{'fixture':10} {'operation':12} {'before cyc':>12} {'after cyc':>12} "
        f"{'delta':>8} {'IPC b':>6} {'IPC a':>6}"
    )
    for stem in STEMS:
        for label in ("open", "one-cell"):
            before, after = (
                perf_isolate(root, "before", stem, label),
                perf_isolate(root, "after", stem, label),
            )
            if not before or not after:
                continue
            print(
                f"{stem:10} {label:12} {before['cycles']:12,.0f} {after['cycles']:12,.0f} "
                f"{100 * (after['cycles'] - before['cycles']) / before['cycles']:+7.2f}% "
                f"{before['instructions'] / before['cycles']:6.2f} "
                f"{after['instructions'] / after['cycles']:6.2f}"
            )
    print()
    print(f"{'fixture':10} {'all-cells leg':22} {'cycles':>16} {'instructions':>16} {'IPC':>6}")
    for stem in STEMS:
        for label in ("all-cells-scan", "all-cells-per-cell"):
            row = perf_isolate(root, "after", stem, label)
            if not row:
                continue
            print(
                f"{stem:10} {label:22} {row['cycles']:16,.0f} {row['instructions']:16,.0f} "
                f"{row['instructions'] / row['cycles']:6.2f}"
            )
    print()


def table_latency(root):
    data = load(root, "latency-summary.json")
    print("== Wall clock, paired, A1 B1 B2 A2 ==")
    print("a1/a2 are the before binary, b1/b2 the after binary. The floor is the")
    print("same binary against itself in the same window: a1 vs a2, b1 vs b2.\n")
    print(
        f"{'case':42} {'a1 p50':>11} {'b1 p50':>11} {'b2 p50':>11} {'a2 p50':>11} "
        f"{'b1/a1':>8} {'b2/a2':>8} {'A/A':>8} {'B/B':>8}"
    )
    floors_aa, floors_bb, deltas = [], [], []
    for case in sorted(data["cells"]):
        rounds = data["cells"][case]
        if not all(key in rounds for key in ("b1", "b2")):
            continue
        has_a = all(key in rounds for key in ("a1", "a2"))
        a1 = rounds["a1"]["p50"] if has_a else float("nan")
        a2 = rounds["a2"]["p50"] if has_a else float("nan")
        b1, b2 = rounds["b1"]["p50"], rounds["b2"]["p50"]
        first = 100 * (b1 - a1) / a1 if has_a else float("nan")
        second = 100 * (b2 - a2) / a2 if has_a else float("nan")
        aa = 100 * abs(a2 - a1) / a1 if has_a else float("nan")
        bb = 100 * abs(b2 - b1) / b1
        if has_a:
            floors_aa.append(aa)
            deltas.extend([first, second])
        floors_bb.append(bb)
        print(
            f"{case:42} {a1:11,.0f} {b1:11,.0f} {b2:11,.0f} {a2:11,.0f} "
            f"{first:+7.2f}% {second:+7.2f}% {aa:7.2f}% {bb:7.2f}%"
        )
    if floors_aa:
        print(
            f"\nA/A floor over {len(floors_aa)} paired cells: "
            f"p50 {statistics.median(floors_aa):.2f}%, mean {statistics.fmean(floors_aa):.2f}%, "
            f"max {max(floors_aa):.2f}%"
        )
    print(
        f"B/B floor over {len(floors_bb)} cells: "
        f"p50 {statistics.median(floors_bb):.2f}%, mean {statistics.fmean(floors_bb):.2f}%, "
        f"max {max(floors_bb):.2f}%"
    )
    if deltas:
        triggers = [d for d in deltas if d > 5.0]
        print(
            f"paired deltas: {len(deltas)} comparisons, "
            f"{sum(1 for d in deltas if d < 0)} improve, "
            f"{len(triggers)} above the +5% review trigger"
        )
    print()


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    if "--fold" in sys.argv:
        fold(sys.argv[sys.argv.index("--fold") + 1], root)
        print("folded")
        return
    table_counters(root)
    table_walk(root)
    table_callgrind(root)
    table_perf(root)
    table_latency(root)


if __name__ == "__main__":
    main()
