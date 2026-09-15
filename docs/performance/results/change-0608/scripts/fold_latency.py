#!/usr/bin/env python3
"""Fold the A1 B1 B2 A2 wall-clock rounds into p50/mean/p95/p99 and the floor."""
import json
import os
import statistics
import sys

PRETTY = {"flagship": "ConditionalFormattingSamples.xls",
          "cv": "WithCustomViews.xls", "54016": "54016.xls"}


def quant(values, q):
    values = sorted(values)
    index = min(len(values) - 1, max(0, int(round(q * (len(values) - 1)))))
    return values[index]


def load_summary(path):
    """Re-read a previously folded summary, so the packet recomputes its own
    tables without the raw per-round captures (24 files, 11 MB, not retained)."""
    with open(path) as handle:
        raw = json.load(handle)
    return {tuple(key.split("|")): value for key, value in raw.items()}


def main(directory, out_json):
    if directory.endswith(".json"):
        cells = load_summary(directory)
        out_json = None
        render(cells, out_json)
        return
    cells = {}
    for name in sorted(os.listdir(directory)):
        if not name.startswith("lat-") or not name.endswith(".json"):
            continue
        stem = name[4:-5]
        bname, rnd, fixture, mode = stem.split("-", 3)
        with open(os.path.join(directory, name)) as handle:
            doc = json.load(handle)
        ns = doc["elapsed_samples_ns"]
        cells.setdefault((bname, fixture, mode), {})[rnd] = dict(
            n=len(ns), p50=statistics.median(ns), mean=statistics.fmean(ns),
            p95=quant(ns, 0.95), p99=quant(ns, 0.99), sha=doc["binary"]["sha256"])
    render(cells, out_json)


def render(cells, out_json):
    print("=" * 96)
    print("Paired wall clock, A1 B1 B2 A2; A is the base, B is a measurement scaffold")
    print("=" * 96)
    print(f"{'scaffold':<9} {'fixture':<34} {'mode':<13} {'a1 p50':>9} {'b1 p50':>9} "
          f"{'b2 p50':>9} {'a2 p50':>9} {'dir1':>8} {'dir2':>8} {'A/A':>8} {'B/B':>8}")
    floors = []
    for key in sorted(cells):
        bname, fixture, mode = key
        r = cells[key]
        if set(r) != {"a1", "b1", "b2", "a2"}:
            continue
        d1 = 100.0 * (r["b1"]["p50"] - r["a1"]["p50"]) / r["a1"]["p50"]
        d2 = 100.0 * (r["b2"]["p50"] - r["a2"]["p50"]) / r["a2"]["p50"]
        aa = 100.0 * (r["a2"]["p50"] - r["a1"]["p50"]) / r["a1"]["p50"]
        bb = 100.0 * (r["b2"]["p50"] - r["b1"]["p50"]) / r["b1"]["p50"]
        floors += [aa, bb]
        print(f"{bname:<9} {PRETTY[fixture]:<34} {mode:<13} {r['a1']['p50']:>9,.0f} "
              f"{r['b1']['p50']:>9,.0f} {r['b2']['p50']:>9,.0f} {r['a2']['p50']:>9,.0f} "
              f"{d1:>7.2f}% {d2:>7.2f}% {aa:>7.2f}% {bb:>7.2f}%")
    print()
    print(f"Measured same-binary floor across all {len(floors)} comparisons in this window: "
          f"{min(floors):.2f}% to {max(floors):.2f}% at p50")
    print()
    print("Tails, per round (ns):")
    print(f"{'scaffold':<9} {'fixture':<34} {'mode':<13} {'round':<5} {'n':>5} "
          f"{'p50':>9} {'mean':>9} {'p95':>9} {'p99':>10}")
    for key in sorted(cells):
        for rnd in ("a1", "b1", "b2", "a2"):
            r = cells[key].get(rnd)
            if r:
                print(f"{key[0]:<9} {PRETTY[key[1]]:<34} {key[2]:<13} {rnd:<5} {r['n']:>5} "
                      f"{r['p50']:>9,.0f} {r['mean']:>9,.0f} {r['p95']:>9,.0f} {r['p99']:>10,.0f}")
    if out_json:
        with open(out_json, "w") as handle:
            json.dump({f"{k[0]}|{k[1]}|{k[2]}": v for k, v in cells.items()}, handle, indent=1)
            handle.write("\n")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
