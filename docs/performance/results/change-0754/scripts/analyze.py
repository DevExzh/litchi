#!/usr/bin/env python3
"""Change 0754 ABBA analysis (the 0747/0750 method): per-process p50/p95/mean,
per-arm median of process p50s, paired after/before ratios (slot 1 with 2,
slot 4 with 3 in every round) with a percentile bootstrap CI, output-digest
parity, and each process's whole-process instructions and cycles from
`perf stat` (corpus construction and the harness's untimed checks included)."""
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
    counters = {}
    perf = path[:-5] + '.perf'
    if os.path.exists(perf):
        for line in open(perf):
            fields = line.strip().split(',')
            if len(fields) > 2 and fields[2] in ('instructions', 'cycles') and fields[0].isdigit():
                counters[fields[2]] = int(fields[0])
    sink = result.get('sink') or {}
    records.setdefault(label, []).append({
        'round': rnd, 'slot': slot, 'arm': arm, 'case': result['case'],
        'shape': result.get('corpus', {}).get('shape'),
        'p50': e['p50'], 'p95': e['p95'], 'mean': e['mean'], 'n': len(e['samples']),
        'output_sha256': result.get('output_sha256') or sink.get('sha256'),
        'instructions': counters.get('instructions'), 'cycles': counters.get('cycles'),
    })

def bootstrap(ratios, iterations=20000, seed=754):
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
    pairs, instruction_pairs, cycle_pairs = [], [], []
    for rnd in sorted({r['round'] for r in rows}):
        for a_slot, b_slot in ((1, 2), (4, 3)):
            a, b = by.get((rnd, a_slot)), by.get((rnd, b_slot))
            if a and b:
                pairs.append(b['p50'] / a['p50'])
                if a['instructions'] and b['instructions']:
                    instruction_pairs.append(b['instructions'] / a['instructions'])
                    cycle_pairs.append(b['cycles'] / a['cycles'])
    lo, hi = bootstrap(pairs)
    digests = {arm: sorted({str(r['output_sha256']) for r in arms[arm]}) for arm in 'AB'}
    def med(arm, key):
        values = [r[key] for r in arms[arm] if r[key] is not None]
        return statistics.median(values) if values else None
    summary[label] = {
        'case': rows[0]['case'], 'shape': rows[0]['shape'],
        'processes': {arm: len(arms[arm]) for arm in 'AB'},
        'samples_per_process': sorted({r['n'] for r in rows}),
        'per_process': sorted(({k: r[k] for k in ('round', 'slot', 'arm', 'p50', 'p95', 'mean', 'instructions', 'cycles')} for r in rows), key=lambda r: (r['round'], r['slot'])),
        'median_p50_ns': {arm: med(arm, 'p50') for arm in 'AB'},
        'median_p95_ns': {arm: med(arm, 'p95') for arm in 'AB'},
        'median_mean_ns': {arm: med(arm, 'mean') for arm in 'AB'},
        'median_process_instructions': {arm: med(arm, 'instructions') for arm in 'AB'},
        'median_process_cycles': {arm: med(arm, 'cycles') for arm in 'AB'},
        'paired_ratios_after_over_before': pairs,
        'paired_ratio_median': statistics.median(pairs),
        'paired_ratio_bootstrap_95': [lo, hi],
        'paired_process_instruction_ratio_median': statistics.median(instruction_pairs) if instruction_pairs else None,
        'paired_process_cycle_ratio_median': statistics.median(cycle_pairs) if cycle_pairs else None,
        'output_sha256': digests,
        'output_identical_across_arms': digests['A'] == digests['B'] and len(digests['A']) == 1,
    }
json.dump(summary, open(os.path.join(OUT, 'analysis.json'), 'w'), indent=2, sort_keys=True)
print(f"{'label':<20} {'shape':<11} {'A p50 ms':>9} {'B p50 ms':>9} {'ratio':>7} {'95% CI':>17} {'A p95':>8} {'B p95':>8} {'proc instr':>10} {'proc cyc':>9} digest")
for label, s in summary.items():
    a, b = s['median_p50_ns']['A'] / 1e6, s['median_p50_ns']['B'] / 1e6
    lo, hi = s['paired_ratio_bootstrap_95']
    ir = s['paired_process_instruction_ratio_median']; cr = s['paired_process_cycle_ratio_median']
    print(f"{label:<20} {str(s['shape']):<11} {a:>9.4f} {b:>9.4f} {s['paired_ratio_median']:>7.4f} [{lo:.4f}, {hi:.4f}] {s['median_p95_ns']['A']/1e6:>8.4f} {s['median_p95_ns']['B']/1e6:>8.4f} {ir if ir is None else round(ir,4):>10} {cr if cr is None else round(cr,4):>9} {'same' if s['output_identical_across_arms'] else 'DIFF ' + str(s['output_sha256'])}")
