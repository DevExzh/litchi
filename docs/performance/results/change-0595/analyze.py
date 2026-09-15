#!/usr/bin/env python3
"""Recomputes every table change 0595 cites, from this directory alone.

    python3 -B docs/performance/results/change-0595/analyze.py \
      docs/performance/results/change-0595

Standard library only. Three inputs, all retained here:

  counters/<leg>-<stem>-<mode>-<op>.json   logical counters per operation
  callgrind/ann-<leg>-<stem>-<op>-s{small,large}.txt
                                           callgrind_annotate self cost, the
                                           isolation pair change 0574 defined
  latency/<round>/<stem>/<mode>-<op>.json  the A1 B1 B2 A2 wall-clock rounds

The callgrind figures are differences between a large-sample and a small-sample
child divided by the extra operations, so everything that runs once per child --
the harness's SHA-256 of the input, its eager oracle projection -- cancels.
"""

import json
import os
import re
import statistics
import sys

# stem: (small, large) sample counts, per operation, matching capture_callgrind.sh
SAMPLES = {
    ("flagship", "open"): (20, 220),
    ("flagship", "one-cell"): (20, 220),
    ("cv", "open"): (20, 120),
    ("cv", "one-cell"): (20, 120),
    ("54016", "open"): (10, 60),
    ("54016", "one-cell"): (5, 25),
}
STEMS = ["flagship", "cv", "54016"]
OPS = ["open", "list", "one-cell"]
MODES = ["owned-readat", "file-source"]
LINE = re.compile(r"^\s*([\d,]+) \([\s\d.]+%\)\s+(.*?)(?: \[[^\]]*\])?$")


def annotate(path):
    """Self Ir per symbol, plus PROGRAM TOTALS, from one callgrind_annotate."""
    totals, symbols = None, {}
    body = False
    for line in open(path, encoding="utf-8", errors="replace"):
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


def isolate(root, leg, stem, op):
    """One operation's Ir: (large - small) / (large samples - small samples)."""
    small, large = SAMPLES[(stem, op)]
    paths = [
        os.path.join(root, "callgrind", f"ann-{leg}-{stem}-{op}-s{label}.txt")
        for label in ("small", "large")
    ]
    if not all(os.path.exists(path) for path in paths):
        return None, {}
    (ts, ss), (tl, sl) = annotate(paths[0]), annotate(paths[1])
    span = large - small
    whole = (tl - ts) / span
    per = {}
    for name in set(ss) | set(sl):
        per[name] = (sl.get(name, 0) - ss.get(name, 0)) / span
    return whole, per


def counters(root):
    """Prefers the folded summary; falls back to the raw capture directory."""
    folded = os.path.join(root, "counters-summary.json")
    if os.path.exists(folded):
        with open(folded, encoding="utf-8") as handle:
            return [
                (row["fixture"], row["mode"], row["operation"], row["cell"])
                for row in json.load(handle)["cells"]
            ]
    return counters_raw(root)


def counters_raw(root):
    rows = []
    for stem in STEMS:
        for mode in MODES:
            for op in OPS:
                cell = {}
                for leg in ("before", "after"):
                    path = os.path.join(
                        root, "counters", f"{leg}-{stem}-{mode}-{op}.json"
                    )
                    if not os.path.exists(path):
                        continue
                    with open(path, encoding="utf-8") as handle:
                        data = json.load(handle)
                    first = data["records"][0]["metrics"]
                    for record in data["records"]:
                        for key in (
                            "read_calls",
                            "read_bytes",
                            "version_calls",
                            "len_calls",
                            "seek_calls",
                        ):
                            assert record["metrics"][key] == first[key], (
                                f"{leg}/{stem}/{mode}/{op}: {key} varies between samples"
                            )
                    cell[leg] = first
                    cell[f"{leg}_cell"] = data["semantic_oracle"][
                        "source_implementation_projection"
                    ]
                if len(cell) >= 4:
                    rows.append((stem, mode, op, cell))
    return rows


def fold(root):
    """Folds the two bulky raw captures into the summaries this record retains.

    The per-sample arrays are megabytes and carry nothing the tables use, so
    only the deterministic counters, the oracle projection, the binary hash and
    the per-round percentiles are kept.
    """
    rows = counters_raw(root)
    payload = {
        "note": "one representative record per cell; every sample in a cell "
        "carries identical logical counters, which counters_raw asserts",
        "cells": [
            {"fixture": stem, "mode": mode, "operation": op, "cell": cell}
            for stem, mode, op, cell in rows
        ],
    }
    with open(os.path.join(root, "counters-summary.json"), "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, sort_keys=True)
        handle.write("\n")

    base = os.path.join(root, "latency")
    summary = {"rounds": {}}
    for round_name in sorted(os.listdir(base)):
        for stem in STEMS:
            directory = os.path.join(base, round_name, stem)
            if not os.path.isdir(directory):
                continue
            for name in sorted(os.listdir(directory)):
                if not name.endswith(".json"):
                    continue
                with open(os.path.join(directory, name), encoding="utf-8") as handle:
                    data = json.load(handle)
                stats = percentiles(data["elapsed_samples_ns"])
                stats["binary_sha256"] = data["binary"]["sha256"]
                stats["warmups"] = data["warmups"]
                key = f"{stem}|{data['mode']}|{data['operation']}"
                summary["rounds"].setdefault(round_name, {})[key] = stats
    with open(os.path.join(root, "latency-summary.json"), "w", encoding="utf-8") as handle:
        json.dump(summary, handle, indent=2, sort_keys=True)
        handle.write("\n")
    print("folded counters-summary.json and latency-summary.json")


def perf_counts(path):
    counts = {}
    if not os.path.exists(path):
        return counts
    for line in open(path, encoding="utf-8"):
        parts = line.strip().split(",")
        if len(parts) < 3 or not parts[0] or parts[0].startswith("#"):
            continue
        try:
            counts[parts[2]] = float(parts[0])
        except ValueError:
            continue
    return counts


def perf(root):
    """Hardware counters per operation, by the same isolation differencing."""
    rows = []
    for stem in STEMS:
        for op in ("open", "one-cell"):
            cell = {}
            for leg in ("before", "after"):
                small = perf_counts(
                    os.path.join(root, "perf", f"perf-{leg}-{stem}-{op}-s100.csv")
                )
                large = perf_counts(
                    os.path.join(root, "perf", f"perf-{leg}-{stem}-{op}-s1100.csv")
                )
                if not small or not large:
                    continue
                cell[leg] = {
                    key: (large[key] - small[key]) / 1000.0
                    for key in small
                    if key in large
                }
            if len(cell) == 2:
                rows.append((stem, op, cell))
    return rows


def percentiles(values):
    ordered = sorted(values)
    def at(fraction):
        index = min(len(ordered) - 1, max(0, int(round(fraction * (len(ordered) - 1)))))
        return ordered[index]
    return {
        "n": len(ordered),
        "p50": statistics.median(ordered),
        "mean": statistics.fmean(ordered),
        "p95": at(0.95),
        "p99": at(0.99),
    }


def latency(root):
    cells = {}
    folded = os.path.join(root, "latency-summary.json")
    if os.path.exists(folded):
        with open(folded, encoding="utf-8") as handle:
            summary = json.load(handle)
        for round_name, entries in summary["rounds"].items():
            for key, stats in entries.items():
                stem, mode, op = key.split("|")
                cells.setdefault((stem, mode, op), {})[round_name] = stats
        return cells
    base = os.path.join(root, "latency")
    if not os.path.isdir(base):
        return cells
    for round_name in sorted(os.listdir(base)):
        for stem in STEMS:
            directory = os.path.join(base, round_name, stem)
            if not os.path.isdir(directory):
                continue
            for name in sorted(os.listdir(directory)):
                if not name.endswith(".json"):
                    continue
                with open(os.path.join(directory, name), encoding="utf-8") as handle:
                    data = json.load(handle)
                key = (stem, data["mode"], data["operation"])
                cells.setdefault(key, {})[round_name] = percentiles(
                    data["elapsed_samples_ns"]
                )
    return cells


def main():
    argv = [value for value in sys.argv[1:] if value != "--fold"]
    root = argv[0] if argv else os.path.dirname(os.path.abspath(__file__))
    if "--fold" in sys.argv:
        fold(root)
        return

    print("=" * 96)
    print("1. Logical counters per operation (the control: no I/O moved)")
    print("=" * 96)
    print(
        f"{'fixture':10} {'mode':13} {'op':9} "
        f"{'reads':>7} {'read bytes':>12} {'version':>8} {'len':>4} {'seek':>5}  identical"
    )
    all_identical = True
    for stem, mode, op, cell in counters(root):
        before, after = cell["before"], cell["after"]
        keys = ("read_calls", "read_bytes", "version_calls", "len_calls", "seek_calls")
        same = all(before[key] == after[key] for key in keys)
        same_cell = cell["before_cell"] == cell["after_cell"]
        all_identical = all_identical and same and same_cell
        print(
            f"{stem:10} {mode:13} {op:9} "
            f"{before['read_calls']:7d} {before['read_bytes']:12d} "
            f"{before['version_calls']:8d} {before['len_calls']:4d} "
            f"{before['seek_calls']:5d}  "
            f"{'yes' if same else 'NO'} / oracle {'yes' if same_cell else 'NO'}"
        )
    print(f"\nevery counter and every oracle projection identical: {all_identical}")

    print()
    print("=" * 96)
    print("2. Instructions per operation, callgrind isolation pairs")
    print("=" * 96)
    print(f"{'fixture':10} {'op':9} {'before Ir':>14} {'after Ir':>14} {'delta':>10}")
    wholes = {}
    for stem in STEMS:
        for op in ("open", "one-cell"):
            if (stem, op) not in SAMPLES:
                continue
            before, before_syms = isolate(root, "before", stem, op)
            after, after_syms = isolate(root, "after", stem, op)
            if before is None or after is None:
                continue
            wholes[(stem, op)] = (before, before_syms, after, after_syms)
            print(
                f"{stem:10} {op:9} {before:14,.0f} {after:14,.0f} "
                f"{(after - before) / before * 100:9.2f}%"
            )

    print()
    print("=" * 96)
    print("3. Where the instructions went, self Ir per operation (movers only)")
    print("=" * 96)
    for (stem, op), (whole_b, before_syms, whole_a, after_syms) in wholes.items():
        print(f"\n-- {stem}/{op}   {whole_b:,.0f} -> {whole_a:,.0f} Ir")
        moves = []
        for name in set(before_syms) | set(after_syms):
            b, a = before_syms.get(name, 0.0), after_syms.get(name, 0.0)
            if abs(a - b) >= max(2000.0, 0.002 * whole_b):
                moves.append((a - b, name, b, a))
        for delta, name, b, a in sorted(moves):
            short = name.split(":")[-1]
            print(f"   {b:13,.0f} -> {a:13,.0f}  {delta:+13,.0f}  {short[:82]}")

    rows = perf(root)
    if rows:
        print()
        print("=" * 96)
        print("4. Hardware counters per operation, perf stat isolation pairs")
        print("=" * 96)
        print(
            f"{'fixture':10} {'op':9} {'cycles B':>12} {'cycles A':>12} {'d cyc':>8} "
            f"{'instr B':>13} {'instr A':>13} {'d ins':>8} {'IPC B':>6} {'IPC A':>6} "
            f"{'miss B':>8} {'miss A':>8}"
        )
        for stem, op, cell in rows:
            b, a = cell["before"], cell["after"]
            print(
                f"{stem:10} {op:9} {b['cycles']:12,.0f} {a['cycles']:12,.0f} "
                f"{(a['cycles'] - b['cycles']) / b['cycles'] * 100:7.2f}% "
                f"{b['instructions']:13,.0f} {a['instructions']:13,.0f} "
                f"{(a['instructions'] - b['instructions']) / b['instructions'] * 100:7.2f}% "
                f"{b['instructions'] / b['cycles']:6.2f} "
                f"{a['instructions'] / a['cycles']:6.2f} "
                f"{b['branch-misses']:8,.0f} {a['branch-misses']:8,.0f}"
            )
        print()
        print("Native counters, so rep movsb is one instruction and SHA-256 uses the")
        print("hardware extension; the callgrind table above counts neither that way.")
        print("Both legs ran in the same window on a shared host: see quiescence.log.")

    cells = latency(root)
    if cells:
        print()
        print("=" * 96)
        print("5. Wall clock, A1 B1 B2 A2, nanoseconds")
        print("=" * 96)
        header = (
            f"{'fixture':10} {'mode':13} {'op':9} "
            f"{'A p50':>9} {'B p50':>9} {'dir1':>8} {'dir2':>8} "
            f"{'A/A':>7} {'B/B':>7} {'B mean':>9} {'B p95':>9} {'B p99':>9}"
        )
        print(header)
        for key in sorted(cells):
            rounds = cells[key]
            if not all(name in rounds for name in ("a1", "a2", "b1", "b2")):
                continue
            a1, a2 = rounds["a1"]["p50"], rounds["a2"]["p50"]
            b1, b2 = rounds["b1"]["p50"], rounds["b2"]["p50"]
            stem, mode, op = key
            print(
                f"{stem:10} {mode:13} {op:9} "
                f"{(a1 + a2) / 2:9.0f} {(b1 + b2) / 2:9.0f} "
                f"{(b1 - a1) / a1 * 100:7.2f}% {(b2 - a2) / a2 * 100:7.2f}% "
                f"{(a2 - a1) / a1 * 100:6.2f}% {(b2 - b1) / b1 * 100:6.2f}% "
                f"{(rounds['b1']['mean'] + rounds['b2']['mean']) / 2:9.0f} "
                f"{(rounds['b1']['p95'] + rounds['b2']['p95']) / 2:9.0f} "
                f"{(rounds['b1']['p99'] + rounds['b2']['p99']) / 2:9.0f}"
            )
        print()
        print("A p50 and B p50 are the mean of the two rounds of that leg.")
        print("dir1 = b1 against a1, dir2 = b2 against a2: negative is faster after.")
        print("A/A = a2 against a1 and B/B = b2 against b1 are the same binary against")
        print("itself across the window, and are the measured floor.")

        print()
        print("=" * 96)
        print("6. Tails per round, nanoseconds (owned-readat), so one noisy round is visible")
        print("=" * 96)
        print(
            f"{'fixture':10} {'op':9} {'round':6} {'n':>5} "
            f"{'p50':>10} {'mean':>10} {'p95':>10} {'p99':>10}"
        )
        for key in sorted(cells):
            if key[1] != "owned-readat":
                continue
            for round_name in ("a1", "b1", "b2", "a2"):
                stats = cells[key].get(round_name)
                if not stats:
                    continue
                print(
                    f"{key[0]:10} {key[2]:9} {round_name:6} {stats['n']:5d} "
                    f"{stats['p50']:10.0f} {stats['mean']:10.0f} "
                    f"{stats['p95']:10.0f} {stats['p99']:10.0f}"
                )
        return
        print("dir1 = b1 against a1, dir2 = b2 against a2: negative is faster after.")
        print("A/A = a2 against a1 and B/B = b2 against b1 are the same binary against")
        print("itself across the window, and are the measured floor.")


if __name__ == "__main__":
    main()
