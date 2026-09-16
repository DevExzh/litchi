import sys, csv
S = sys.argv[1]
def read(tag):
    vals = {}
    for line in open(f'{S}/perf-{tag}.txt'):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        parts = line.split(',')
        if len(parts) >= 3 and parts[2] in ('cycles', 'instructions'):
            vals[parts[2]] = int(parts[0])
    return vals
legs = {t: read(t) for t in ('A1', 'B1', 'B2', 'A2')}
print('whole-child perf stat (4 selectors in one process, --warmup 1 --samples 10, taskset -c 11)')
print(f'{"leg":>6s} {"cycles":>20s} {"instructions":>20s}')
for t in ('A1', 'B1', 'B2', 'A2'):
    print(f'{t:>6s} {legs[t]["cycles"]:>20,} {legs[t]["instructions"]:>20,}')
for ev in ('cycles', 'instructions'):
    b = (legs['A1'][ev] + legs['A2'][ev]) / 2
    a = (legs['B1'][ev] + legs['B2'][ev]) / 2
    aa = 100 * (legs['A2'][ev] - legs['A1'][ev]) / legs['A1'][ev]
    bb = 100 * (legs['B2'][ev] - legs['B1'][ev]) / legs['B1'][ev]
    print(f'{ev}: before mean {b:,.0f}  after mean {a:,.0f}  delta {100*(a-b)/b:+.2f}%   A/A {aa:+.2f}%  B/B {bb:+.2f}%')
