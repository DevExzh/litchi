#!/usr/bin/env python3
"""ABBA-interleaved, core-pinned runs of registered harness selectors (change 0757).

Each (selector, writer shape) runs 8 processes in the order A B B A A B B A
(A = before harness, B = after harness), with its own --samples/--warmup.
Binary paths and report paths have equal lengths in both legs. Raw JSON
reports are kept; the summary reports per-process p50/p95/mean of the timed
samples, the median of process p50s per arm, the paired after/before p50
ratios (one per adjacent AB/BA pair) and a percentile bootstrap CI over them.

usage: abba_harness.py OUT_DIR CORE [selector/shape,... [LABEL]]
"""
import json, os, random, statistics, subprocess, sys

OUT, CORE = sys.argv[1], sys.argv[2]
SCRATCH = '/home/zhuhe/code/litchi-worktrees/scratch/0757'
BIN = {'A': f'{SCRATCH}/bin/A/lpb', 'B': f'{SCRATCH}/bin/B/lpb'}
ORDER = ['A', 'B', 'B', 'A', 'A', 'B', 'B', 'A']
# (selector, writer shape, warmup, samples)
CASES = [
    ('xls_fresh_write_to', 'tiny', 100, 1000),
    ('xls_fresh_write_to', 'large', 10, 100),
    ('xls_fresh_write_to', 'payload-heavy', 5, 40),
    ('xls_semantic_one_edit_save', 'tiny', 40, 400),
    ('xls_semantic_one_edit_save', 'large', 10, 100),
    ('doc_fresh_write_to', 'tiny', 100, 1000),
    ('doc_fresh_write_to', 'large', 20, 200),
    ('doc_fresh_write_to', 'payload-heavy', 5, 40),
    ('ppt_fresh_write_to', 'tiny', 100, 1000),
    ('ppt_fresh_write_to', 'large', 20, 200),
    ('ppt_fresh_write_to', 'payload-heavy', 5, 40),
]


def pct(values, q):
    values = sorted(values)
    k = (len(values) - 1) * q
    lo, hi = int(k), min(int(k) + 1, len(values) - 1)
    return values[lo] + (values[hi] - values[lo]) * (k - lo)


def bootstrap(ratios, iterations=10000, seed=757):
    rng = random.Random(seed)
    boots = sorted(statistics.median([rng.choice(ratios) for _ in ratios]) for _ in range(iterations))
    return (boots[int(0.025 * iterations)], boots[int(0.975 * iterations) - 1])


def main():
    selected = sys.argv[3].split(',') if len(sys.argv) > 3 else None
    os.makedirs(os.path.join(OUT, 'raw'), exist_ok=True)
    summary = []
    for selector, shape, warmup, samples in CASES:
        name = f'{selector}/{shape}'
        if selected and name not in selected:
            continue
        processes = []
        for index, leg in enumerate(ORDER):
            raw = os.path.join(OUT, 'raw', f'{selector}-{shape}-{index}-{leg}.json')
            command = ['taskset', '-c', CORE, BIN[leg], '--case', selector, '--writer-shape', shape,
                       '--samples', str(samples), '--warmup', str(warmup), '--json', raw]
            subprocess.run(command, capture_output=True, text=True, check=True)
            with open(raw) as handle:
                report = json.load(handle)
            result = [r for r in report['results'] if r['case'] == selector][0]
            values = result['elapsed_ns']['samples']
            processes.append({
                'index': index, 'leg': leg, 'raw': os.path.relpath(raw, OUT),
                'p50': pct(values, 0.5), 'p95': pct(values, 0.95), 'mean': statistics.fmean(values),
                'n': len(values), 'corpus': result['corpus'].get('name'),
                'archive_sha256': result['corpus'].get('archive_sha256'),
            })
        a = [p for p in processes if p['leg'] == 'A']
        b = [p for p in processes if p['leg'] == 'B']
        ratios, mean_ratios = [], []
        for i in range(0, len(ORDER), 2):
            first, second = processes[i], processes[i + 1]
            before = first if first['leg'] == 'A' else second
            after = second if second['leg'] == 'B' else first
            ratios.append(after['p50'] / before['p50'])
            mean_ratios.append(after['mean'] / before['mean'])
        lo, hi = bootstrap(ratios)
        record = {
            'case': selector, 'shape': shape, 'warmup': warmup, 'samples': samples,
            'corpus': processes[0]['corpus'],
            'archive_sha256_before': sorted({p['archive_sha256'] for p in a}),
            'archive_sha256_after': sorted({p['archive_sha256'] for p in b}),
            'before_median_p50': statistics.median(p['p50'] for p in a),
            'after_median_p50': statistics.median(p['p50'] for p in b),
            'before_median_p95': statistics.median(p['p95'] for p in a),
            'after_median_p95': statistics.median(p['p95'] for p in b),
            'before_median_mean': statistics.median(p['mean'] for p in a),
            'after_median_mean': statistics.median(p['mean'] for p in b),
            'paired_p50_ratios': ratios, 'paired_mean_ratios': mean_ratios,
            'median_paired_p50_ratio': statistics.median(ratios),
            'bootstrap_95ci_p50_ratio': [lo, hi],
            'processes': processes,
        }
        summary.append(record)
        print(f"{name:42s} before {record['before_median_p50']/1e6:9.4f} ms  after "
              f"{record['after_median_p50']/1e6:9.4f} ms  ratio {record['median_paired_p50_ratio']:.3f} "
              f"CI [{lo:.3f}, {hi:.3f}]  same-output={record['archive_sha256_before'] == record['archive_sha256_after']}", flush=True)
    label = sys.argv[4] if len(sys.argv) > 4 else 'partial'
    out = 'summary.json' if not selected else f'summary-{label}.json'
    with open(os.path.join(OUT, out), 'w') as handle:
        json.dump(summary, handle, indent=1)


if __name__ == '__main__':
    main()
