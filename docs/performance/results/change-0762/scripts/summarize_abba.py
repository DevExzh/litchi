#!/usr/bin/env python3
"""Summarize an ABBA campaign of harness JSON reports (changes 0762/0763).

Pairs within each round: slot 1 (before) with slot 2 (after), slot 4
(before) with slot 3 (after). Reports per-process p50/p95/mean, the median
of process p50s per leg, the median paired p50 change, a percentile
bootstrap 95% interval over the paired changes (20,000 resamples, seed 762),
every paired p50/p95/mean comparison beyond 5%, and output-digest identity
within each leg and across the legs.
"""
import glob, json, os, random, re, statistics, sys

def load(directory):
    runs = {}
    for path in sorted(glob.glob(os.path.join(directory, '*.json'))):
        m = re.match(r'(.+)-r(\d+)-s(\d)-(before|after)\.json$', os.path.basename(path))
        if not m:
            continue
        name, rnd, slot, leg = m.group(1), int(m.group(2)), int(m.group(3)), m.group(4)
        data = json.load(open(path))
        for result in data['results']:
            key = (name, result['corpus']['name'])
            e = result['elapsed_ns']
            runs.setdefault(key, []).append({
                'round': rnd, 'slot': slot, 'leg': leg, 'p50': e['p50'] / 1e6,
                'p95': e['p95'] / 1e6, 'mean': e['mean'] / 1e6,
                'samples': len(e['samples']), 'sha': result.get('output_sha256'),
                'binary_sha256': data['binary_identity']['binary_sha256'], 'file': os.path.basename(path)})
    return runs

def bootstrap(values, seed=762, resamples=20000):
    rng = random.Random(seed)
    stats = sorted(statistics.median(rng.choices(values, k=len(values))) for _ in range(resamples))
    return stats[int(0.025 * resamples)], stats[int(0.975 * resamples) - 1]

def main():
    directory = sys.argv[1]
    runs = load(directory)
    summary, flags = [], []
    for (name, corpus), rows in sorted(runs.items()):
        before = [r for r in rows if r['leg'] == 'before']
        after = [r for r in rows if r['leg'] == 'after']
        pairs = []
        for rnd in sorted({r['round'] for r in rows}):
            get = {(r['slot']): r for r in rows if r['round'] == rnd}
            for b, a in ((1, 2), (4, 3)):
                if b in get and a in get:
                    pairs.append((get[b], get[a]))
        changes = [a['p50'] / b['p50'] - 1 for b, a in pairs]
        lo, hi = bootstrap(changes)
        for b, a in pairs:
            for metric in ('p50', 'p95', 'mean'):
                change = a[metric] / b[metric] - 1
                if abs(change) > 0.05:
                    flags.append({'case': name, 'corpus': corpus, 'round': b['round'], 'metric': metric,
                                  'before_ms': round(b[metric], 4), 'after_ms': round(a[metric], 4),
                                  'change': round(change, 4),
                                  'direction': 'adverse' if change > 0 else 'favourable'})
        digests = sorted({r['sha'] or 'none' for r in rows})
        before_digests = sorted({r['sha'] or 'none' for r in before})
        after_digests = sorted({r['sha'] or 'none' for r in after})
        summary.append({
            'case': name, 'corpus': corpus, 'processes_before': len(before), 'processes_after': len(after),
            'samples_per_process': sorted({r['samples'] for r in rows}),
            'before_p50_median_ms': round(statistics.median(r['p50'] for r in before), 4),
            'after_p50_median_ms': round(statistics.median(r['p50'] for r in after), 4),
            'before_p95_median_ms': round(statistics.median(r['p95'] for r in before), 4),
            'after_p95_median_ms': round(statistics.median(r['p95'] for r in after), 4),
            'before_mean_median_ms': round(statistics.median(r['mean'] for r in before), 4),
            'after_mean_median_ms': round(statistics.median(r['mean'] for r in after), 4),
            'paired_p50_changes': [round(c, 4) for c in changes],
            'median_paired_p50_change': round(statistics.median(changes), 4),
            'bootstrap_95ci': [round(lo, 4), round(hi, 4)],
            'output_sha256': digests, 'outputs_identical_across_legs': len(digests) == 1,
            'before_output_sha256': before_digests, 'after_output_sha256': after_digests,
            'each_leg_deterministic': len(before_digests) == 1 and len(after_digests) == 1,
            'binaries': sorted({(r['leg'], r['binary_sha256']) for r in rows}),
            'processes': sorted(({k: r[k] for k in ('round', 'slot', 'leg', 'p50', 'p95', 'mean', 'file')} for r in rows),
                                key=lambda r: (r['round'], r['slot'])),
        })
    json.dump({'summary': summary, 'flags_over_5_percent': flags}, open(os.path.join(directory, 'summary.json'), 'w'), indent=1)
    print(f"{'case':30} {'corpus':40} {'before':>10} {'after':>10} {'paired':>8} {'95% CI':>18} per-leg-det same-across")
    for s in summary:
        print(f"{s['case']:30} {s['corpus']:40} {s['before_p50_median_ms']:10.3f} {s['after_p50_median_ms']:10.3f} "
              f"{s['median_paired_p50_change']:+8.2%} [{s['bootstrap_95ci'][0]:+.2%}, {s['bootstrap_95ci'][1]:+.2%}] {s['each_leg_deterministic']} {s['outputs_identical_across_legs']}")
    adverse = [f for f in flags if f['direction'] == 'adverse']
    print(f"flags over 5%: {len(flags)} ({len(adverse)} adverse)")
    for f in adverse:
        print('  ADVERSE', f)

if __name__ == '__main__':
    main()
