"""Percentiles and paired deltas for change 0636's timing legs."""
import sys, statistics, glob, os

def load(path):
    return sorted(int(line) for line in open(path) if line.strip())

def pct(values, q):
    if not values:
        return 0
    index = min(len(values) - 1, max(0, int(round(q * (len(values) - 1)))))
    return values[index]

def describe(values):
    return (pct(values, 0.50), statistics.mean(values), pct(values, 0.95), pct(values, 0.99))

rows = []
base = sys.argv[1]
labels = sorted({os.path.basename(p).split('-', 1)[1].rsplit('-', 1)[0]
                 for p in glob.glob(os.path.join(base, 't-*-A1.txt'))})
print(f"{'scenario':38} {'leg':4} {'p50 us':>10} {'mean us':>10} {'p95 us':>10} {'p99 us':>10}")
for label in labels:
    legs = {}
    for leg in ('A1', 'B1', 'B2', 'A2'):
        path = os.path.join(base, f't-{label}-{leg}.txt')
        legs[leg] = load(path)
        p50, mean, p95, p99 = describe(legs[leg])
        print(f"{label:38} {leg:4} {p50/1000:10.3f} {mean/1000:10.3f} {p95/1000:10.3f} {p99/1000:10.3f}")
    a = sorted(legs['A1'] + legs['A2'])
    b = sorted(legs['B1'] + legs['B2'])
    aa = (pct(legs['A2'], 0.50) - pct(legs['A1'], 0.50)) / pct(legs['A1'], 0.50) * 100
    bb = (pct(legs['B2'], 0.50) - pct(legs['B1'], 0.50)) / pct(legs['B1'], 0.50) * 100
    delta50 = (pct(b, 0.50) - pct(a, 0.50)) / pct(a, 0.50) * 100
    delta95 = (pct(b, 0.95) - pct(a, 0.95)) / pct(a, 0.95) * 100
    delta99 = (pct(b, 0.99) - pct(a, 0.99)) / pct(a, 0.99) * 100
    print(f"{label:38} {'A/A':4} floor p50 {aa:+7.2f}%   B/B {bb:+7.2f}%")
    print(f"{label:38} {'A→B':4} p50 {delta50:+7.2f}%  p95 {delta95:+7.2f}%  p99 {delta99:+7.2f}%")
    print()
