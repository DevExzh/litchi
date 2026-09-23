#!/usr/bin/env python3
"""ABBA-interleaved, core-pinned runs of registered harness selectors.

Each selector runs 8 processes in the order A B B A A B B A (A = before
harness, B = after harness), one selector per process, with its own
--samples/--warmup. Raw JSON reports are kept; the summary reports per-process
p50/p95/mean of `elapsed_ns`, the median of process p50s per arm, the paired
after/before p50 ratios and a percentile bootstrap CI over them.

usage: abba_harness.py OUT_DIR CORE
"""
import json, os, random, statistics, subprocess, sys

OUT = sys.argv[1]
CORE = sys.argv[2]
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
BIN = {'A': os.environ.get('HARNESS_BEFORE', os.path.join(ROOT, 'bin', 'harness-before')), 'B': os.path.join(ROOT, 'bin', 'harness-after')}
ORDER = ['A', 'B', 'B', 'A', 'A', 'B', 'B', 'A']
FIXTURE = os.path.join(ROOT, 'fixtures', '54016.xls')

LARGE = ['--writer-shape', 'large']
# (selector, warmup, samples, extra args)
CASES = [
    ('xls_semantic_one_edit_save', 5, 40, LARGE),
    ('xls_semantic_noop_edit_save', 5, 40, LARGE),
    ('xls_semantic_open', 5, 40, LARGE),
    ('xls_semantic_full_cell_scan', 5, 40, LARGE),
    ('xls_numeric_eager_number_edit_save', 5, 30, []),
    ('xls_numeric_source_backed_number_edit_save', 5, 30, []),
    ('xls_numeric_plan_only_number_edit_save', 5, 30, []),
    ('xls_numeric_eager_rk_mulrk_edit_save', 5, 60, []),
    ('xls_numeric_source_backed_rk_mulrk_edit_save', 5, 60, []),
    ('xls_numeric_plan_only_rk_mulrk_edit_save', 5, 60, []),
    ('xls_comments_eager_edit_save', 5, 30, []),
    ('xls_comments_source_backed_edit_save', 5, 30, []),
    ('xls_comments_eager_batch_edit_save', 5, 30, []),
    ('xls_comments_source_backed_batch_edit_save', 5, 30, []),
    ('xls_visibility_eager_edit_save', 5, 60, []),
    ('xls_visibility_source_backed_edit_save', 5, 60, []),
    ('xls_visibility_eager_batch_edit_save', 5, 60, []),
    ('xls_visibility_source_backed_batch_edit_save', 5, 60, []),
    ('xls_owned_source_control_open_one_cell', 5, 60, ['--ole2-file', FIXTURE]),
]


def pct(values, q):
    values = sorted(values)
    k = (len(values) - 1) * q
    lo, hi = int(k), min(int(k) + 1, len(values) - 1)
    return values[lo] + (values[hi] - values[lo]) * (k - lo)


def bootstrap(ratios, iterations=10000, seed=746):
    rng = random.Random(seed)
    boots = sorted(statistics.median([rng.choice(ratios) for _ in ratios]) for _ in range(iterations))
    return (boots[int(0.025 * iterations)], boots[int(0.975 * iterations) - 1])


def main():
    selected = sys.argv[3].split(',') if len(sys.argv) > 3 else None
    os.makedirs(os.path.join(OUT, 'raw'), exist_ok=True)
    summary = []
    for selector, warmup, samples, extra in CASES:
        if selected and selector not in selected:
            continue
        processes = []
        for index, leg in enumerate(ORDER):
            raw = os.path.join(OUT, 'raw', f'{selector}-{index}-{leg}.json')
            command = ['taskset', '-c', CORE, BIN[leg], '--case', selector, '--samples', str(samples),
                       '--warmup', str(warmup), '--json', raw, *extra]
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
            'case': selector, 'warmup': warmup, 'samples': samples, 'extra': extra,
            'corpus': processes[0]['corpus'],
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
        print(f"{selector:48s} before {record['before_median_p50']/1e6:9.4f} ms  after "
              f"{record['after_median_p50']/1e6:9.4f} ms  ratio {record['median_paired_p50_ratio']:.3f} "
              f"CI [{lo:.3f}, {hi:.3f}]", flush=True)
    name = 'summary.json' if not selected else 'summary-partial.json'
    with open(os.path.join(OUT, name), 'w') as handle:
        json.dump(summary, handle, indent=1)


if __name__ == '__main__':
    main()
