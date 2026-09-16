"""Decomposes the measured DOC saving into a per-byte term (the three removed
complete scans) and a fixed term (the removed composed CFB reopen and the
metadata observations that went with it), by ordinary least squares on

    saving(fixture) = 3 * c * bytes + delta_fixed
"""
import sys, statistics

def load(path):
    out = {}
    for ordinal, line in enumerate(open(path)):
        if line.startswith('mode\t'):
            continue
        mode, fixture, size, cyc, ins, *_ = line.rstrip('\n').split('\t')
        out[(mode, fixture, ordinal)] = (int(size), float(cyc))
    return out

a1, b1, b2, a2 = (load(p) for p in sys.argv[1:5])
mode = sys.argv[5]
scans = int(sys.argv[6])
rows = []
for key in a1:
    if key[0] != mode:
        continue
    size = a1[key][0]
    before = (a1[key][1] + a2[key][1]) / 2
    after = (b1[key][1] + b2[key][1]) / 2
    rows.append((size, before - after, before, after))
rows.sort()
n = len(rows)
mx = sum(r[0] for r in rows) / n
my = sum(r[1] for r in rows) / n
slope = sum((r[0] - mx) * (r[1] - my) for r in rows) / sum((r[0] - mx) ** 2 for r in rows)
intercept = my - slope * mx
print(f"{mode}: saving = {slope:.6f} * bytes + {intercept:,.0f}   ({n} fixtures, OLS)")
print(f"  per-scan constant = {slope / scans:.4f} cycles per byte per scan "
      f"({scans} scans removed); change 0644's DOC OLS constant is 2.1633")
print(f"  fixed term = {intercept:,.0f} cycles per operation")
print("fixture_bytes\tmeasured_saving\tfitted_saving\tresidual%")
res = []
for size, saving, before, after in rows:
    fit = slope * size + intercept
    r = (fit - saving) / saving * 100
    res.append(r)
    print(f"{size}\t{saving:.0f}\t{fit:.0f}\t{r:+.2f}")
print(f"# fit residual median {statistics.median(res):+.2f}%, range {min(res):+.2f}%..{max(res):+.2f}%")
