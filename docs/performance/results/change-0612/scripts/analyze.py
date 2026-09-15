#!/usr/bin/env python3
"""Fold change 0612's isolation pairs into per-operation figures.

Change 0574's method, as 0576, 0584, 0595 and 0608 used it: the same child runs
at a small and a large sample count, and differencing the two profiles and
dividing by the extra operations cancels everything that runs once per child.
Both the self-cost annotation (`ann-*`) and the inclusive annotation (`inc-*`)
are folded, because the size of the globals pass is the *inclusive* cost of
`parse_globals` and the size of the term this change moves is the inclusive
cost of the CFB range reader inside it.

Pure standard library.

Usage: analyze.py <callgrind-dir> [<perf-dir> ...]
"""
import os
import re
import statistics
import sys
from collections import defaultdict

# stem -> op -> (small samples, large samples); must match capture_callgrind.sh
SAMPLES = {
    "flagship": {"open": (20, 220), "one-cell": (20, 220)},
    "cv": {"open": (20, 120), "one-cell": (20, 120)},
    "54016": {"open": (10, 60), "one-cell": (10, 60)},
}
PRETTY = {
    "flagship": "ConditionalFormattingSamples.xls",
    "cv": "WithCustomViews.xls",
    "54016": "54016.xls",
}
ORDER = ("flagship", "cv", "54016")
ROW = re.compile(r"^\s*([\d,]+) \(\s*[\d.]+%\)\s+(.*?)\s*$")

# Inclusive symbols the record cites.
GLOBALS_PASS = "litchi_xls::workbook::source::parse_globals"
RANGE_READ = "litchi_cfb::shared::SharedOleFile::read_stream_range_hinted"


def parse(path):
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
    out = {}
    pat = re.compile(
        rf"^{kind}-(.+?)-(flagship|cv|54016)-(open|one-cell)-s(small|large)\.txt$"
    )
    groups = defaultdict(dict)
    for name in os.listdir(directory):
        m = pat.match(name)
        if m:
            groups[(m.group(1), m.group(2), m.group(3))][m.group(4)] = os.path.join(
                directory, name
            )
    for key, files in sorted(groups.items()):
        if set(files) != {"small", "large"}:
            continue
        leg, stem, op = key
        small_n, large_n = SAMPLES[stem][op]
        delta_n = large_n - small_n
        st, ss = parse(files["small"])
        lt, ls = parse(files["large"])
        syms = {}
        for sym in set(ss) | set(ls):
            syms[sym] = (ls.get(sym, 0) - ss.get(sym, 0)) / delta_n
        out[key] = ((lt - st) / delta_n, syms)
    return out


def find(syms, needle):
    """Inclusive value of the one symbol whose demangled name contains `needle`."""
    best = 0.0
    for name, value in syms.items():
        if needle in name:
            best = max(best, value)
    return best


def main(argv):
    cg = argv[0]
    incl = fold(cg, "inc")
    self_ = fold(cg, "ann")
    legs = [leg for leg in ("base", "skip", "skipcur", "cursor") if any(k[0] == leg for k in incl)]

    print("=" * 78)
    print("1. Per-operation instructions, callgrind isolation pairs (owned source)")
    print("=" * 78)
    print(f"{'fixture':<34} {'op':<9} " + " ".join(f"{leg:>13}" for leg in legs) + f" {'base-skip':>13} {'share':>8}")
    for stem in ORDER:
        for op in ("open", "one-cell"):
            cells, vals = [], {}
            for leg in legs:
                v = incl.get((leg, stem, op))
                vals[leg] = v[0] if v else None
                cells.append(f"{v[0]:>13,.0f}" if v else f"{'-':>13}")
            if vals.get("base") and vals.get("skip"):
                d = vals["base"] - vals["skip"]
                tail = f" {d:>13,.0f} {100.0 * d / vals['base']:>7.2f}%"
            else:
                tail = ""
            if any(c.strip() != "-" for c in cells):
                print(f"{PRETTY[stem]:<34} {op:<9} " + " ".join(cells) + tail)

    print()
    print("=" * 78)
    print("2. The globals pass, INCLUSIVE Ir per operation and its share")
    print("=" * 78)
    print(
        f"{'fixture':<34} {'op':<9} {'leg':<6} {'parse_globals':>14} {'share':>8} "
        f"{'range read':>13} {'share':>8}"
    )
    for stem in ORDER:
        for op in ("open", "one-cell"):
            for leg in legs:
                v = incl.get((leg, stem, op))
                if not v:
                    continue
                total, syms = v
                g = find(syms, "parse_globals")
                r = find(syms, "read_stream_range_hinted")
                print(
                    f"{PRETTY[stem]:<34} {op:<9} {leg:<6} {g:>14,.0f} "
                    f"{100.0 * g / total:>7.2f}% {r:>13,.0f} {100.0 * r / total:>7.2f}%"
                )

    print()
    print("=" * 78)
    print("3. Self Ir per operation, symbols above 0.3% of the base open")
    print("=" * 78)
    for stem in ORDER:
        for op in ("open", "one-cell"):
            base = self_.get(("base", stem, op))
            if not base:
                continue
            total, syms = base
            rows = [(s, v) for s, v in syms.items() if abs(v) > 0.003 * total]
            rows.sort(key=lambda r: -r[1])
            print(f"\n-- {PRETTY[stem]}  {op}  (base total {total:,.0f} Ir)")
            print(f"{'symbol':<58}" + "".join(f"{leg:>13}" for leg in legs) + f"{'delta':>13}")
            for sym, _v in rows:
                cells = ""
                vals = {}
                for leg in legs:
                    lv = self_.get((leg, stem, op))
                    value = lv[1].get(sym, 0.0) if lv else 0.0
                    vals[leg] = value
                    cells += f"{value:>13,.0f}"
                delta = vals.get("base", 0.0) - vals.get("skip", 0.0)
                print(f"{sym[:58]:<58}" + cells + f"{delta:>+13,.0f}")

    for directory in argv[1:]:
        perf(directory, os.path.basename(directory.rstrip("/")))


def perf(directory, label):
    print()
    print("=" * 78)
    print(f"4. Native perf stat, isolation pairs, median of repetitions [{label}]")
    print("=" * 78)
    pat = re.compile(
        r"^perf-(.+?)-(flagship|cv|54016)-(open|one-cell)-s(\d+)-r(\d+)\.csv$"
    )
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
    print(
        f"{'fixture':<34} {'op':<9} {'leg':<6} {'cycles':>12} {'instructions':>14} "
        f"{'IPC':>6} {'reps':>5}"
    )
    for key, samples in sorted(runs.items()):
        if len(samples) < 2:
            continue
        small, large = min(samples), max(samples)
        delta = large - small
        out = {}
        for event in ("cycles", "instructions", "branch-misses", "task-clock"):
            lo = statistics.median([v.get(event, 0.0) for v in samples[small]])
            hi = statistics.median([v.get(event, 0.0) for v in samples[large]])
            out[event] = (hi - lo) / delta
        per_op[key] = out
        leg, stem, op = key
        ipc = out["instructions"] / out["cycles"] if out["cycles"] else 0.0
        print(
            f"{PRETTY[stem]:<34} {op:<9} {leg:<6} {out['cycles']:>12,.0f} "
            f"{out['instructions']:>14,.0f} {ipc:>6.2f} {len(samples[large]):>5}"
        )

    print()
    print("Pairs (positive = the first leg costs more):")
    print(
        f"{'fixture':<34} {'op':<9} {'comparison':<14} {'cycles':>13} {'share':>8} "
        f"{'Ir':>13} {'share':>8}"
    )
    pairs = [
        ("base", "skip"),
        ("base", "skipcur"),
        ("base", "cursor"),
        ("aa", "bb"),
        ("base", "aa"),
        ("bb", "base"),
    ]
    for stem in ORDER:
        for op in ("open", "one-cell"):
            for left, right in pairs:
                a = per_op.get((left, stem, op))
                b = per_op.get((right, stem, op))
                if not a or not b:
                    continue
                dc = a["cycles"] - b["cycles"]
                di = a["instructions"] - b["instructions"]
                print(
                    f"{PRETTY[stem]:<34} {op:<9} {left + '-' + right:<14} {dc:>13,.0f} "
                    f"{100.0 * dc / a['cycles']:>7.2f}% {di:>13,.0f} "
                    f"{100.0 * di / a['instructions']:>7.2f}%"
                )


if __name__ == "__main__":
    main(sys.argv[1:])
