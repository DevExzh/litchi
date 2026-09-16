#!/usr/bin/env python3
"""Change 0644: per-scan cost and the predicted saving of each design variant.

Variants differ only in which empty-splice identity calls take one scan:
  B3 -- only inside ensure_current, never the DOC open (open = 6 - C's 2 = 4);
  B2 -- every call except the last of its operation (open = 3);
  B1 -- every call (open = 2).

Run from the packet root:  python3 scripts/predict.py
Reads perf/perfstat-legA.tsv, perf/perfstat-legB.tsv and perf/perfstat-cfb-index.tsv.

Two estimators are reported for the per-scan constant, because they answer
different questions and the record cites both:
  * OLS over every fixture of the format  -- the estimate;
  * a two-point interpolation through the smallest and largest fixture -- the
    check that the relation is linear rather than an artefact of the middle.
Residuals are (predicted - measured) / measured, stated in that direction.
"""
import csv
import statistics
import sys

def load(path):
    with open(path) as handle:
        return list(csv.DictReader(handle, delimiter="\t"))

def ols(points):
    """Least squares over (bytes, cycles). Returns (intercept, slope)."""
    n = len(points)
    mx = sum(p[0] for p in points) / n
    my = sum(p[1] for p in points) / n
    sxy = sum((p[0] - mx) * (p[1] - my) for p in points)
    sxx = sum((p[0] - mx) ** 2 for p in points)
    slope = sxy / sxx
    return my - slope * mx, slope

def endpoints(points):
    lo, hi = points[0], points[-1]
    slope = (hi[1] - lo[1]) / (hi[0] - lo[0])
    return lo[1] - slope * lo[0], slope

def residuals(points, intercept, slope):
    return [((intercept + slope * b) - c) / c * 100 for b, c in points]

legA = load("perf/perfstat-legA.tsv")
legB = load("perf/perfstat-legB.tsv")
index = {r["fixture"]: float(r["cycles_per_op"]) for r in load("perf/perfstat-cfb-index.tsv")}

fits = {}
print("## Per-scan cost (measured)\n")
print("mode\testimator\tintercept\tslope_per_byte\tscans\tcycles_per_byte_per_scan\tresidual_median\tresidual_min\tresidual_max")
for mode, scans in (("doc-open", 6), ("ppt-open", 2)):
    pts = sorted((int(r["bytes"]), float(r["cycles_per_op"])) for r in legA if r["mode"] == mode)
    for name, fit in (("ols", ols), ("two-point", endpoints)):
        ic, sl = fit(pts)
        res = residuals(pts, ic, sl)
        if name == "ols":
            fits[mode] = (ic, sl, sl / scans, len(pts))
        print(f"{mode}\t{name}\t{ic:,.0f}\t{sl:.4f}\t{scans}\t{sl/scans:.4f}"
              f"\t{statistics.median(res):+.2f}%\t{min(res):+.2f}%\t{max(res):+.2f}%")

per_doc = fits["doc-open"][2]
per_ppt = fits["ppt-open"][2]
print(f"\n# OLS agreement between the two formats: {abs(per_doc - per_ppt) / per_ppt * 100:.2f}%")
print("# 0609 fitted 217,274 + 12.90b over 6 scans -> 2.1500; 0589 measured 2.048 (2.029-2.115)")

print("\n## A/A floor, cycles per operation, leg B against leg A\n")
deltas = sorted((float(b["cycles_per_op"]) - float(a["cycles_per_op"])) / float(a["cycles_per_op"]) * 100
                for a, b in zip(legA, legB))
absd = sorted(abs(d) for d in deltas)
print(f"cells\t{len(deltas)}\nmedian\t{statistics.median(deltas):+.2f}%\n"
      f"min\t{min(deltas):+.2f}%\nmax\t{max(deltas):+.2f}%\n"
      f"p95_abs\t{absd[int(0.95 * len(absd)) - 1]:.2f}%")

print("\n## Predicted DOC open after each variant (modelled: measured before - scans x per-scan x bytes)\n")
print("fixture\tbytes\tmeasured_before\tB3C_after(6->4)\tB3C_delta\tB2C_after(6->3)\tB2C_delta\tB1C_after(6->2)\tB1C_delta")
rows = []
for r in sorted((r for r in legA if r["mode"] == "doc-open"), key=lambda r: int(r["bytes"])):
    b, before = int(r["bytes"]), float(r["cycles_per_op"])
    out = [r["fixture"], f"{b}", f"{before:,.0f}"]
    deltas_row = []
    for k in (2, 3, 4):
        save = k * per_doc * b
        out += [f"{before - save:,.0f}", f"{-save / before * 100:+.1f}%"]
        deltas_row.append(-save / before * 100)
    rows.append(deltas_row)
    print("\t".join(out))
for label, i in (("B3+C (6->4), the recommendation; B omitted from the open, so this is also C alone", 0),
                 ("B2+C (6->3)", 1), ("B1+C (6->2)", 2)):
    col = [row[i] for row in rows]
    print(f"# {label}: median {statistics.median(col):+.1f}%, range {min(col):+.1f}%..{max(col):+.1f}%")

print("\n## Predicted further saving on the readback (resolve calls ensure_current twice)\n")
print("fixture\tbytes\tB3_or_B2 (8->6 scans)\tB1 (8->4 scans)")
for r in sorted((r for r in legA if r["mode"] == "doc-open"), key=lambda r: int(r["bytes"])):
    b = int(r["bytes"])
    print(f'{r["fixture"]}\t{b}\t{2 * per_doc * b:,.0f}\t{4 * per_doc * b:,.0f}')

print("\n## One CFB index parse (measured), the term Options D and E would buy\n")
doc_idx = {k: v for k, v in index.items() if k.endswith(".doc")}
print(f"across the 8 DOC fixtures\t{min(doc_idx.values()):,.0f}..{max(doc_idx.values()):,.0f} cycles")
print(f"across all 38 fixtures\t{min(index.values()):,.0f}..{max(index.values()):,.0f} cycles")
opens = {r["fixture"]: float(r["cycles_per_op"]) for r in legA if r["mode"] == "doc-open"}
share = sorted(doc_idx[f] / opens[f] * 100 for f in opens)
print(f"one parse as a share of the DOC open\t{share[0]:.2f}%..{share[-1]:.2f}%")
two = sorted(2 * doc_idx[f] / opens[f] * 100 for f in opens)
print(f"the two reopens B+C retains\t{two[0]:.2f}%..{two[-1]:.2f}%")
