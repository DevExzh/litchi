#!/usr/bin/env python3
"""Fold change 0608's callgrind isolation pairs into per-operation figures.

Change 0574's method, as 0576, 0584 and 0595 used it: the same child is run at
a small and a large sample count; differencing the two profiles and dividing by
the extra operations cancels everything that runs once per child.  Both the
self-cost annotation (`ann-*`) and the inclusive annotation (`inc-*`) are folded,
because this record needs the inclusive cost of `scan_shared_string_records` --
the walk's share of the open -- as well as the self costs inside it.

Pure standard library.  Usage: analyze.py <callgrind-dir> [<perf-dir>]
"""
import os
import re
import sys
from collections import defaultdict

# stem -> op -> (small samples, large samples)
SAMPLES = {
    "flagship": {"open": (20, 220), "one-cell": (20, 220)},
    "cv": {"open": (20, 120), "one-cell": (20, 120)},
    "54016": {"open": (10, 60), "one-cell": (5, 25)},
}
PRETTY = {
    "flagship": "ConditionalFormattingSamples.xls",
    "cv": "WithCustomViews.xls",
    "54016": "54016.xls",
}
ROW = re.compile(r"^\s*([\d,]+) \(\s*[\d.]+%\)\s+(.*?)\s*$")


def parse(path):
    """total Ir, and {symbol: Ir} from one callgrind_annotate listing."""
    total = None
    syms = {}
    for line in open(path, encoding="utf-8", errors="replace"):
        m = ROW.match(line)
        if not m:
            continue
        value = int(m.group(1).replace(",", ""))
        name = m.group(2)
        if name == "PROGRAM TOTALS":
            total = value
            continue
        if total is None:
            continue
        name = re.sub(r"\s*\[[^\]]*\]\s*$", "", name)
        name = name.split(":", 1)[1] if ":" in name else name
        syms[name] = syms.get(name, 0) + value
    return total, syms


def fold(directory, kind):
    """{(leg, stem, op): (per-op total, {symbol: per-op Ir})}"""
    out = {}
    pat = re.compile(rf"^{kind}-(.+?)-(flagship|cv|54016)-(open|one-cell)-s(small|large)\.txt$")
    groups = defaultdict(dict)
    for name in os.listdir(directory):
        m = pat.match(name)
        if m:
            groups[(m.group(1), m.group(2), m.group(3))][m.group(4)] = os.path.join(directory, name)
    for key, files in sorted(groups.items()):
        if set(files) != {"small", "large"}:
            continue
        leg, stem, op = key
        small_n, large_n = SAMPLES[stem][op]
        delta_n = large_n - small_n
        st, ss = parse(files["small"])
        lt, ls = parse(files["large"])
        per_op = (lt - st) / delta_n
        syms = {}
        for sym in set(ss) | set(ls):
            syms[sym] = (ls.get(sym, 0) - ss.get(sym, 0)) / delta_n
        out[key] = (per_op, syms)
    return out


def main(argv):
    cg = argv[0]
    incl = fold(cg, "inc")
    self_ = fold(cg, "ann")

    legs = sorted({k[0] for k in incl})
    print("=" * 78)
    print("Per-operation instructions (callgrind isolation pairs)")
    print("=" * 78)
    print(f"{'fixture':<34} {'op':<9} " + " ".join(f"{leg:>14}" for leg in legs))
    for stem in ("flagship", "cv", "54016"):
        for op in ("open", "one-cell"):
            cells = []
            for leg in legs:
                v = incl.get((leg, stem, op))
                cells.append(f"{v[0]:>14,.0f}" if v else f"{'-':>14}")
            if any(c.strip() != "-" for c in cells):
                print(f"{PRETTY[stem]:<34} {op:<9} " + " ".join(cells))

    print()
    print("=" * 78)
    print("scan_shared_string_records, INCLUSIVE Ir per operation and its share")
    print("=" * 78)
    key = "litchi_xls::records::scan_shared_string_records"
    print(f"{'fixture':<34} {'op':<9} {'leg':<9} {'inclusive':>12} {'total':>12} {'share':>8}")
    for stem in ("flagship", "cv", "54016"):
        for op in ("open", "one-cell"):
            for leg in legs:
                v = incl.get((leg, stem, op))
                if not v:
                    continue
                total, syms = v
                val = syms.get(key, 0.0)
                print(f"{PRETTY[stem]:<34} {op:<9} {leg:<9} {val:>12,.0f} {total:>12,.0f} "
                      f"{100.0 * val / total:>7.2f}%")

    print()
    print("=" * 78)
    print("Deltas against the base leg (inclusive totals)")
    print("=" * 78)
    for stem in ("flagship", "cv", "54016"):
        for op in ("open", "one-cell"):
            base = incl.get(("base", stem, op))
            if not base:
                continue
            for leg in legs:
                if leg == "base":
                    continue
                v = incl.get((leg, stem, op))
                if not v:
                    continue
                d = v[0] - base[0]
                print(f"{PRETTY[stem]:<34} {op:<9} base - {leg:<8} "
                      f"{-d:>12,.0f} Ir  ({-100.0 * d / base[0]:>6.2f}% of the operation)")

    print()
    print("=" * 78)
    print("Self Ir per operation, symbols whose base value exceeds 0.2% of the open")
    print("=" * 78)
    for stem in ("flagship", "cv", "54016"):
        for op in ("open", "one-cell"):
            base = self_.get(("base", stem, op))
            if not base:
                continue
            total, syms = base
            rows = [(s, v) for s, v in syms.items() if abs(v) > 0.002 * total]
            rows.sort(key=lambda r: -r[1])
            print(f"\n-- {PRETTY[stem]}  {op}  (total {total:,.0f} Ir)")
            header = f"{'symbol':<62}" + "".join(f"{leg:>13}" for leg in legs)
            print(header)
            for sym, _v in rows:
                cells = ""
                for leg in legs:
                    lv = self_.get((leg, stem, op))
                    cells += f"{lv[1].get(sym, 0.0):>13,.0f}" if lv else f"{'-':>13}"
                print(f"{sym[:62]:<62}" + cells)

    if len(argv) > 1:
        perf(argv[1], legs)


def perf(directory, _legs):
    """Native counters: median over repetitions of each isolation pair leg.

    Each (leg, fixture, op, sample count) was run REPS times; the median of the
    repetitions is taken per leg before the pair is differenced, so one
    scheduling excursion on this shared host cannot set the result.
    """
    import statistics
    print()
    print("=" * 78)
    print("Native perf stat, isolation pairs, median of repetitions")
    print("=" * 78)
    pat = re.compile(r"^perf-(.+?)-(flagship|cv|54016)-(open|one-cell)-s(\d+)-r(\d+)\.csv$")
    runs = defaultdict(lambda: defaultdict(list))
    for name in os.listdir(directory):
        m = pat.match(name)
        if not m:
            continue
        vals = {}
        for line in open(os.path.join(directory, name)):
            parts = line.strip().split(",")
            if len(parts) > 2 and parts[0] and parts[0][0].isdigit():
                try:
                    vals[parts[2]] = float(parts[0])
                except ValueError:
                    pass
        runs[(m.group(1), m.group(2), m.group(3))][int(m.group(4))].append(vals)

    per_op = {}
    print(f"{'fixture':<34} {'op':<9} {'leg':<9} {'cycles':>12} {'instructions':>14} "
          f"{'IPC':>6} {'br.miss':>9} {'reps':>5}")
    for key, samples in sorted(runs.items()):
        if len(samples) < 2:
            continue
        small, large = min(samples), max(samples)
        delta = large - small
        out = {}
        for event in ("cycles", "instructions", "branch-misses"):
            lo = statistics.median([v.get(event, 0.0) for v in samples[small]])
            hi = statistics.median([v.get(event, 0.0) for v in samples[large]])
            out[event] = (hi - lo) / delta
        per_op[key] = out
        leg, stem, op = key
        ipc = out["instructions"] / out["cycles"] if out["cycles"] else 0.0
        print(f"{PRETTY[stem]:<34} {op:<9} {leg:<9} {out['cycles']:>12,.0f} "
              f"{out['instructions']:>14,.0f} {ipc:>6.2f} {out['branch-misses']:>9,.0f} "
              f"{len(samples[large]):>5}")

    print()
    print("Deltas against the base leg:")
    print(f"{'fixture':<34} {'op':<9} {'comparison':<18} {'cycles saved':>13} {'share':>8} "
          f"{'Ir saved':>13} {'share':>8}")
    for (leg, stem, op), out in sorted(per_op.items()):
        if leg == "base":
            continue
        base = per_op.get(("base", stem, op))
        if not base:
            continue
        dc = base["cycles"] - out["cycles"]
        di = base["instructions"] - out["instructions"]
        print(f"{PRETTY[stem]:<34} {op:<9} {'base - ' + leg:<18} {dc:>13,.0f} "
              f"{100.0 * dc / base['cycles']:>7.2f}% {di:>13,.0f} "
              f"{100.0 * di / base['instructions']:>7.2f}%")


if __name__ == "__main__":
    main(sys.argv[1:])
