"""Paired-timing summary: p50, mean, p95 and p99 per leg, the paired delta in
both directions, and the A/A floor from the two before legs and the two after
legs of the same A1 B1 B2 A2 window."""
import sys, os, glob, statistics, collections

def load(path):
    return sorted(int(line) for line in open(path) if line.strip())

def pct(values, q):
    return values[min(len(values) - 1, int(q * len(values)))]

def stat(values):
    return (statistics.median(values), statistics.mean(values), pct(values, 0.95), pct(values, 0.99))

groups = collections.defaultdict(dict)
for path in sorted(glob.glob(os.path.join(sys.argv[1], 'bench-*-A1.txt')) + glob.glob(os.path.join(sys.argv[1], 'bench-*-A2.txt')) + glob.glob(os.path.join(sys.argv[1], 'bench-*-B1.txt')) + glob.glob(os.path.join(sys.argv[1], 'bench-*-B2.txt'))):
    stem = os.path.basename(path)[len('bench-'):-len('.txt')]
    key, leg = stem.rsplit('-', 1)
    groups[key][leg] = load(path)

print("scenario\tn\tbefore_p50\tafter_p50\tdelta_p50%\tbefore_mean\tafter_mean\t"
      "before_p95\tafter_p95\tbefore_p99\tafter_p99\tAA_before%\tAA_after%")
floor = []
for key in sorted(groups):
    legs = groups[key]
    if set(legs) != {'A1', 'B1', 'B2', 'A2'}:
        continue
    before = sorted(legs['A1'] + legs['A2'])
    after = sorted(legs['B1'] + legs['B2'])
    bs, as_ = stat(before), stat(after)
    delta = (as_[0] - bs[0]) / bs[0] * 100
    aa_b = (statistics.median(legs['A2']) - statistics.median(legs['A1'])) / statistics.median(legs['A1']) * 100
    aa_a = (statistics.median(legs['B2']) - statistics.median(legs['B1'])) / statistics.median(legs['B1']) * 100
    floor += [aa_b, aa_a]
    print(f"{key}\t{len(before)}\t{bs[0]:.0f}\t{as_[0]:.0f}\t{delta:+.2f}\t{bs[1]:.0f}\t{as_[1]:.0f}\t"
          f"{bs[2]}\t{as_[2]}\t{bs[3]}\t{as_[3]}\t{aa_b:+.2f}\t{aa_a:+.2f}")
    print(f"#   {key}: after against before {delta:+.2f}%; before against after "
          f"{(bs[0] - as_[0]) / as_[0] * 100:+.2f}%")
absf = sorted(abs(v) for v in floor)
print(f"\n# A/A floor over {len(floor)} same-leg pairs: median {statistics.median(floor):+.2f}%, "
      f"range {min(floor):+.2f}%..{max(floor):+.2f}%, p95 of |delta| {absf[int(0.95 * (len(absf) - 1))]:.2f}%")
