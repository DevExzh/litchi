#!/usr/bin/env python3
"""Change 0750 ABBA analysis (0747 method, unchanged): per-process p50/p95/mean, per-arm median of
process p50s, paired after/before ratios (slot 1 with 2, slot 4 with 3 in every
round) with a percentile bootstrap CI, output-digest parity and, for the
source-backed XLSX cases, per-phase medians."""
import glob, json, os, random, re, statistics, sys

OUT = sys.argv[1]
raw = os.path.join(OUT, 'raw')
records = {}
for path in sorted(glob.glob(os.path.join(raw, '*.json'))):
    name = os.path.basename(path)[:-5]
    m = re.match(r'(.+)-r(\d+)-s(\d+)-([AB])$', name)
    label, rnd, slot, arm = m.group(1), int(m.group(2)), int(m.group(3)), m.group(4)
    doc = json.load(open(path))
    result = doc['results'][0]
    e = result['elapsed_ns']
    phases = {}
    x = result.get('source', {}).get('xlsx_cell_values') if isinstance(result.get('source'), dict) else None
    if isinstance(x, dict):
        for key in ('open_ns', 'plan_ns', 'commit_ns', 'publication_ns'):
            if isinstance(x.get(key), list) and x[key]:
                phases[key] = statistics.median(x[key])
    records.setdefault(label, []).append({
        'round': rnd, 'slot': slot, 'arm': arm, 'case': result['case'],
        'p50': e['p50'], 'p95': e['p95'], 'mean': e['mean'], 'n': len(e['samples']),
        'output_sha256': result.get('output_sha256'), 'phases': phases,
    })

def bootstrap(ratios, iterations=20000, seed=750):
    rng = random.Random(seed)
    stats = []
    for _ in range(iterations):
        sample = [rng.choice(ratios) for _ in ratios]
        stats.append(statistics.median(sample))
    stats.sort()
    return stats[int(0.025 * iterations)], stats[int(0.975 * iterations) - 1]

summary = {}
for label, rows in records.items():
    by = {(r['round'], r['slot']): r for r in rows}
    arms = {arm: [r for r in rows if r['arm'] == arm] for arm in 'AB'}
    pairs = []
    for rnd in sorted({r['round'] for r in rows}):
        for a_slot, b_slot in ((1, 2), (4, 3)):
            a, b = by.get((rnd, a_slot)), by.get((rnd, b_slot))
            if a and b:
                pairs.append(b['p50'] / a['p50'])
    lo, hi = bootstrap(pairs)
    digests = {arm: sorted({str(r['output_sha256']) for r in arms[arm]}) for arm in 'AB'}
    phase_medians = {}
    for arm in 'AB':
        keys = sorted({k for r in arms[arm] for k in r['phases']})
        phase_medians[arm] = {k: statistics.median([r['phases'][k] for r in arms[arm] if k in r['phases']]) for k in keys}
    summary[label] = {
        'case': rows[0]['case'],
        'processes': {arm: len(arms[arm]) for arm in 'AB'},
        'samples_per_process': sorted({r['n'] for r in rows}),
        'per_process': sorted(({k: r[k] for k in ('round', 'slot', 'arm', 'p50', 'p95', 'mean')} for r in rows), key=lambda r: (r['round'], r['slot'])),
        'median_p50_ns': {arm: statistics.median([r['p50'] for r in arms[arm]]) for arm in 'AB'},
        'median_p95_ns': {arm: statistics.median([r['p95'] for r in arms[arm]]) for arm in 'AB'},
        'median_mean_ns': {arm: statistics.median([r['mean'] for r in arms[arm]]) for arm in 'AB'},
        'paired_ratios_after_over_before': pairs,
        'paired_ratio_median': statistics.median(pairs),
        'paired_ratio_bootstrap_95': [lo, hi],
        'output_sha256': digests,
        'output_identical_across_arms': digests['A'] == digests['B'] and len(digests['A']) == 1,
        'phase_medians_ns': phase_medians,
    }
json.dump(summary, open(os.path.join(OUT, 'analysis.json'), 'w'), indent=2, sort_keys=True)
print(f"{'label':<20} {'A p50 ms':>10} {'B p50 ms':>10} {'ratio':>8} {'95% CI':>18} {'A p95':>9} {'B p95':>9} digest")
for label, s in summary.items():
    a, b = s['median_p50_ns']['A'] / 1e6, s['median_p50_ns']['B'] / 1e6
    lo, hi = s['paired_ratio_bootstrap_95']
    print(f"{label:<20} {a:>10.3f} {b:>10.3f} {s['paired_ratio_median']:>8.4f} [{lo:.4f}, {hi:.4f}] {s['median_p95_ns']['A']/1e6:>9.3f} {s['median_p95_ns']['B']/1e6:>9.3f} {'same' if s['output_identical_across_arms'] else 'DIFF'}")
    for arm in 'AB':
        if s['phase_medians_ns'][arm]:
            print('   ', arm, {k: round(v / 1e6, 3) for k, v in s['phase_medians_ns'][arm].items()})
