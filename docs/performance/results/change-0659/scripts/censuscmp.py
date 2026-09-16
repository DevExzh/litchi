"""Per-fixture before/after census deltas: read calls, bytes read as a multiple
of the artifact, and the `len`/`version` observation counts."""
import sys, statistics

def load(path):
    out = {}
    for line in open(path):
        if line.startswith('file\t'):
            continue
        f, b, rc, rb, fa, lc, vc, outcome = line.rstrip('\n').split('\t')
        out[f] = (int(b), int(rc), int(rb), float(fa), int(lc), int(vc), outcome)
    return out

before, after = load(sys.argv[1]), load(sys.argv[2])
print("file\tbytes\toutcome\treads\tread_x\tlen\tversion")
ratios = []
for f, vb in before.items():
    va = after[f]
    assert vb[6] == va[6], f
    print(f"{f.rsplit('/', 1)[1]}\t{vb[0]}\t{vb[6][:34]}\t{vb[1]}->{va[1]}\t{vb[3]:.2f}->{va[3]:.2f}\t"
          f"{vb[4]}->{va[4]}\t{vb[5]}->{va[5]}")
    ratios.append((vb[6].startswith('OK'), vb[3], va[3], vb[1], va[1]))
adm = [r for r in ratios if r[0]]
ref = [r for r in ratios if not r[0]]
for name, group in (("admitted", adm), ("refused", ref)):
    if not group:
        continue
    red = [(1 - a / b) * 100 for _, b, a, _, _ in group]
    print(f"# {name}: {len(group)} fixtures; bytes read "
          f"{min(b for _, b, _, _, _ in group):.2f}x-{max(b for _, b, _, _, _ in group):.2f}x -> "
          f"{min(a for _, _, a, _, _ in group):.2f}x-{max(a for _, _, a, _, _ in group):.2f}x; "
          f"reduction median {statistics.median(red):.1f}%, range {min(red):.1f}%-{max(red):.1f}%")
