#!/usr/bin/env python3
"""Change 0750 audit-probe ABBA analysis: per process, the median of 41 batch
means per case; per arm, the median of the process medians; paired ratios
(slot 2 over 1 and 3 over 4 in every round) with a percentile bootstrap."""
import glob, json, os, random, re, statistics, sys

OUT = sys.argv[1]
rows = {}
for path in sorted(glob.glob(os.path.join(OUT, 'raw', 'probe-*.jsonl'))):
    m = re.search(r'probe-r(\d+)-s(\d+)-([AB])\.jsonl$', path)
    rnd, slot, arm = int(m.group(1)), int(m.group(2)), m.group(3)
    for line in open(path):
        d = json.loads(line)
        v = d['ns_per_audit']
        rows.setdefault(d['case'], []).append({'round': rnd, 'slot': slot, 'arm': arm,
            'median': statistics.median(v), 'p95': sorted(v)[int(0.95 * len(v)) - 1],
            'mean': statistics.mean(v), 'verdict': d['verdict'], 'batches': len(v)})

def bootstrap(values, iterations=20000, seed=750):
    rng = random.Random(seed)
    stats = sorted(statistics.median([rng.choice(values) for _ in values]) for _ in range(iterations))
    return stats[int(0.025 * iterations)], stats[int(0.975 * iterations) - 1]

summary = {}
print(f"{'case':<40} {'A med us':>10} {'B med us':>10} {'ratio':>8} {'95% CI':>18}")
for case, rs in rows.items():
    by = {(r['round'], r['slot']): r for r in rs}
    pairs = [by[(rnd, b)]['median'] / by[(rnd, a)]['median']
             for rnd in sorted({r['round'] for r in rs}) for a, b in ((1, 2), (4, 3))
             if (rnd, a) in by and (rnd, b) in by]
    arms = {arm: [r for r in rs if r['arm'] == arm] for arm in 'AB'}
    med = {arm: statistics.median([r['median'] for r in arms[arm]]) for arm in 'AB'}
    lo, hi = bootstrap(pairs)
    summary[case] = {'per_process': sorted(rs, key=lambda r: (r['round'], r['slot'])),
                     'median_of_process_medians_ns': med, 'paired_ratios': pairs,
                     'paired_ratio_median': statistics.median(pairs), 'paired_ratio_bootstrap_95': [lo, hi],
                     'verdicts': sorted({str(r['verdict']) for r in rs})}
    print(f"{case:<40} {med['A']/1000:>10.1f} {med['B']/1000:>10.1f} {statistics.median(pairs):>8.4f} [{lo:.4f}, {hi:.4f}]")
json.dump(summary, open(os.path.join(OUT, 'analysis.json'), 'w'), indent=2, sort_keys=True)
