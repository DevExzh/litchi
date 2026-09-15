import sys, statistics, glob, os, re, collections
def load(p): return sorted(int(x) for x in open(p) if x.strip())
def pct(v, q):
    idx = min(len(v)-1, max(0, int(round(q*(len(v)-1)))))
    return v[idx]
SC = os.environ['SC']
cases = collections.defaultdict(dict)
for p in glob.glob(SC + '/timing/*-*.txt'):
    b = os.path.basename(p)[:-4]
    m = re.match(r'(.+)-(A1|A2|B1|B2)$', b)
    if not m: continue
    cases[m.group(1)][m.group(2)] = load(p)
hdr = f"{'case/leg':28s} {'n':>5s} {'p50':>9s} {'mean':>9s} {'p95':>9s} {'p99':>9s}"
print(hdr)
for case in sorted(cases):
    legs = cases[case]
    if len(legs) < 4: continue
    for tag in ['A1','A2','B1','B2']:
        v = legs[tag]
        print(f"{case+'/'+tag:28s} {len(v):5d} {pct(v,.50):9,d} {statistics.mean(v):9,.0f} {pct(v,.95):9,d} {pct(v,.99):9,d}")
    A = sorted(legs['A1'] + legs['A2']); B = sorted(legs['B1'] + legs['B2'])
    aa = 100*(pct(legs['A2'],.50)-pct(legs['A1'],.50))/pct(legs['A1'],.50)
    aa99 = 100*(pct(legs['A2'],.99)-pct(legs['A1'],.99))/pct(legs['A1'],.99)
    bb = 100*(pct(legs['B2'],.50)-pct(legs['B1'],.50))/pct(legs['B1'],.50)
    print(f"  A/A floor p50 {aa:+.2f}%  p99 {aa99:+.2f}%   B/B p50 {bb:+.2f}%")
    for q,name in [(.50,'p50'),(.95,'p95'),(.99,'p99')]:
        a, b = pct(A,q), pct(B,q)
        print(f"  pooled {name}: before {a:,d} ns  after {b:,d} ns   after/before {100*(b-a)/a:+.2f}%   before/after {100*(a-b)/b:+.2f}%")
    print()
