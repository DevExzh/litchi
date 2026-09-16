"""change 0635: paired A1 B1 B2 A2 deltas beside the four-run A/A floor."""
import json, os, sys

OUT = sys.argv[1] if len(sys.argv) > 1 else 'out'
STATS = ['p50', 'mean', 'p95', 'p99']

def load(tag):
    rows = {}
    for suffix in ['', '-producer']:
        path = os.path.join(OUT, f'{tag}{suffix}.json')
        if not os.path.exists(path):
            continue
        with open(path) as handle:
            data = json.load(handle)
        for record in data['results']:
            key = (record['case'], record['corpus'].get('shape') or '-')
            rows[key] = record['elapsed_ns']
    return rows

legs = {tag: load(tag) for tag in ['A1', 'B1', 'B2', 'A2', 'F1', 'F2', 'F3', 'F4']}
keys = sorted(set(legs['A1']) & set(legs['B1']))

print('A/A floor in this window, (max-min)/min over the four before-only runs F1..F4')
print(f"{'case':56s} {'shape':14s} " + ' '.join(f'{s:>9s}' for s in STATS))
floor = {}
for key in keys:
    row = []
    for stat in STATS:
        values = [legs[tag][key][stat] for tag in ['F1', 'F2', 'F3', 'F4']]
        row.append((max(values) - min(values)) / min(values) * 100)
    floor[key] = dict(zip(STATS, row))
    print(f"{key[0]:56s} {key[1]:14s} " + ' '.join(f'{v:8.3f}%' for v in row))

print()
print('paired deltas, positive = this change faster; leg 1 is A1 vs B1, leg 2 is A2 vs B2')
for stat in STATS:
    print(f'-- {stat}')
    print(f"{'case':56s} {'shape':14s} {'leg1':>9s} {'leg2':>9s} {'floor':>9s}")
    for key in keys:
        a1 = legs['A1'][key][stat]
        b1 = legs['B1'][key][stat]
        a2 = legs['A2'][key][stat]
        b2 = legs['B2'][key][stat]
        d1 = (a1 - b1) / a1 * 100
        d2 = (a2 - b2) / a2 * 100
        print(f"{key[0]:56s} {key[1]:14s} {d1:8.2f}% {d2:8.2f}% {floor[key][stat]:8.2f}%")
    print()

print('raw p50 nanoseconds')
print(f"{'case':56s} {'shape':14s} " + ' '.join(f'{t:>12s}' for t in ['A1','B1','B2','A2']))
for key in keys:
    print(f"{key[0]:56s} {key[1]:14s} " + ' '.join(
        f"{legs[tag][key]['p50']:12,d}" for tag in ['A1','B1','B2','A2']))
