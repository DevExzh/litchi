"""Change 0659's gate-4 arithmetic.

Reads the four `perf stat` isolation-pair legs (A1 B1 B2 A2), reports per
fixture the before and after cycles, the measured delta, the value change 0644
predicted for the declared variant, and the residual against it; then the A/A
floor from the two before legs and the two after legs.
"""
import sys, os, statistics

# Change 0644's DOC OLS per-scan constant, unrounded.
PER_SCAN = 12.979873 / 6

def load(path):
    # Three `.ppt` basenames repeat across fixture directories, so the row
    # ordinal is part of the key: both legs walk the same list in the same
    # order.
    out = {}
    for ordinal, line in enumerate(open(path)):
        if line.startswith('mode\t'):
            continue
        mode, fixture, size, cyc, ins, cpb, ipb = line.rstrip('\n').split('\t')
        out[(mode, fixture, ordinal)] = (int(size), float(cyc), float(ins))
    return out

def median(values):
    return statistics.median(values) if values else float('nan')

a1, b1, b2, a2 = (load(p) for p in sys.argv[1:5])
scans_removed = {'doc-open': 3, 'ppt-open': 0, 'doc-open-read': 5}

print("mode\tfixture\tbytes\tbefore_cyc\tafter_cyc\tdelta%\tpredicted_after\tresidual%\tAA_before%\tAA_after%")
rows = []
for key in sorted(a1, key=lambda k: (k[0], a1[k][0])):
    mode, fixture, _ordinal = key
    size, cb, _ = a1[key]
    _, cb2, _ = a2[key]
    _, ca, _ = b1[key]
    _, ca2, _ = b2[key]
    before = (cb + cb2) / 2
    after = (ca + ca2) / 2
    delta = (after - before) / before * 100
    k = scans_removed[mode]
    pred = before - k * PER_SCAN * size
    residual = (after - pred) / pred * 100 if pred else float('nan')
    aa_b = (cb2 - cb) / cb * 100
    aa_a = (ca2 - ca) / ca * 100
    rows.append((mode, fixture, size, before, after, delta, pred, residual, aa_b, aa_a))
    print(f"{mode}\t{fixture}\t{size}\t{before:.0f}\t{after:.0f}\t{delta:+.2f}\t{pred:.0f}\t"
          f"{residual:+.2f}\t{aa_b:+.2f}\t{aa_a:+.2f}")

print()
for mode in sorted({r[0] for r in rows}):
    sub = [r for r in rows if r[0] == mode]
    d = [r[5] for r in sub]
    res = [r[7] for r in sub]
    print(f"# {mode}: {len(sub)} fixtures; delta median {median(d):+.2f}%, "
          f"range {min(d):+.2f}%..{max(d):+.2f}%; residual against 0644's prediction "
          f"median {median(res):+.2f}%, range {min(res):+.2f}%..{max(res):+.2f}%")
    worst = max(res, key=abs)
    print(f"#   worst |residual| {worst:+.2f}%  -> gate {'PASS' if abs(worst) <= 5 else 'FAIL'} (+-5%)")
    if mode == 'ppt-open':
        adverse = max(d)
        print(f"#   worst adverse delta {adverse:+.2f}% -> gate {'PASS' if adverse <= 5 else 'FAIL'} (<=+5%)")

floor = [r[8] for r in rows] + [r[9] for r in rows]
absf = sorted(abs(v) for v in floor)
p95 = absf[int(0.95 * (len(absf) - 1))]
print(f"\n# A/A floor over {len(floor)} same-leg pairs: median {median(floor):+.2f}%, "
      f"range {min(floor):+.2f}%..{max(floor):+.2f}%, p95 of |delta| {p95:.2f}%")
