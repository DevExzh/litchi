#!/usr/bin/env python3
"""ABBA-interleaved, core-pinned timing of the change-0746 XLS probe.

Each case (fixture, operation) runs 8 processes per round set in the order
A B B A A B B A (A = before, B = after), each process taking WARMUP untimed
iterations and SAMPLES timed ones. Raw JSON of every process is kept; the
summary reports per-process p50/p95/mean, the median of process p50s per arm,
the paired after/before ratios (pairs in sequence order) and a percentile
bootstrap CI over those paired ratios.

usage: abba_probe.py OUT_DIR CORE
"""
import json, os, random, statistics, subprocess, sys

OUT = sys.argv[1]
CORE = sys.argv[2]
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
BIN = {'A': os.path.join(ROOT, 'bin', 'probe-before'), 'B': os.path.join(ROOT, 'bin', 'probe-after')}
FIX = os.path.join(ROOT, 'fixtures')
ORDER = ['A', 'B', 'B', 'A', 'A', 'B', 'B', 'A']

# (fixture, operation, warmups, samples, metric)
CASES = [
    ('54016.xls', 'open', 3, 20, 'open_ns'),
    ('54016.xls', 'number-plan', 3, 20, 'commit_ns'),
    ('54016.xls', 'number-source-backed', 3, 20, 'commit_ns'),
    ('54016.xls', 'number-generic', 3, 20, 'commit_ns'),
    ('54016.xls', 'string-generic', 3, 20, 'commit_ns'),
    ('54016.xls', 'comments-open', 3, 20, 'open_ns'),
    ('54016.xls', 'visibility-open', 3, 20, 'open_ns'),
    ('54016.xls', 'reader-open', 3, 20, 'open_ns'),
    ('xls-large.xls', 'open', 5, 60, 'open_ns'),
    ('xls-large.xls', 'number-plan', 5, 60, 'commit_ns'),
    ('xls-large.xls', 'number-source-backed', 5, 60, 'commit_ns'),
    ('xls-large.xls', 'number-generic', 5, 60, 'commit_ns'),
    ('xls-large.xls', 'comments-open', 5, 60, 'open_ns'),
    ('xls-large.xls', 'visibility-open', 5, 60, 'open_ns'),
    ('xls-large.xls', 'reader-open', 5, 60, 'open_ns'),
]


def pct(values, q):
    values = sorted(values)
    if not values:
        return None
    k = (len(values) - 1) * q
    lo, hi = int(k), min(int(k) + 1, len(values) - 1)
    return values[lo] + (values[hi] - values[lo]) * (k - lo)


def stats(values):
    return {
        'p50': pct(values, 0.5),
        'p95': pct(values, 0.95),
        'mean': statistics.fmean(values),
        'min': min(values),
        'max': max(values),
        'n': len(values),
    }


def bootstrap(ratios, iterations=10000, seed=746):
    rng = random.Random(seed)
    boots = []
    for _ in range(iterations):
        sample = [rng.choice(ratios) for _ in ratios]
        boots.append(statistics.median(sample))
    boots.sort()
    return (boots[int(0.025 * iterations)], boots[int(0.975 * iterations) - 1])


def main():
    os.makedirs(os.path.join(OUT, 'raw'), exist_ok=True)
    summary = []
    for fixture, operation, warmups, samples, metric in CASES:
        case = f'{fixture}:{operation}'
        processes = []
        for index, leg in enumerate(ORDER):
            raw = os.path.join(OUT, 'raw', f'{fixture}-{operation}-{index}-{leg}.json')
            command = ['taskset', '-c', CORE, BIN[leg], '--input', os.path.join(FIX, fixture),
                       '--operation', operation, '--warmups', str(warmups), '--samples', str(samples)]
            result = subprocess.run(command, capture_output=True, text=True, check=True)
            with open(raw, 'w') as handle:
                handle.write(result.stdout)
            data = json.loads(result.stdout)
            values = data[metric]
            entry = {'index': index, 'leg': leg, 'raw': os.path.relpath(raw, OUT), **stats(values)}
            if metric != 'open_ns' and data.get('open_ns'):
                entry['open_p50'] = pct(data['open_ns'], 0.5)
            entry['published_bytes'] = data.get('published_bytes')
            processes.append(entry)
        a = [p for p in processes if p['leg'] == 'A']
        b = [p for p in processes if p['leg'] == 'B']
        pairs = [(processes[i], processes[i + 1]) for i in range(0, len(ORDER), 2)]
        ratios = []
        for first, second in pairs:
            before = first if first['leg'] == 'A' else second
            after = second if second['leg'] == 'B' else first
            ratios.append(after['p50'] / before['p50'])
        mean_ratios = []
        for first, second in pairs:
            before = first if first['leg'] == 'A' else second
            after = second if second['leg'] == 'B' else first
            mean_ratios.append(after['mean'] / before['mean'])
        lo, hi = bootstrap(ratios)
        record = {
            'case': case,
            'metric': metric,
            'warmups': warmups,
            'samples': samples,
            'before_median_p50': statistics.median(p['p50'] for p in a),
            'after_median_p50': statistics.median(p['p50'] for p in b),
            'before_median_p95': statistics.median(p['p95'] for p in a),
            'after_median_p95': statistics.median(p['p95'] for p in b),
            'before_median_mean': statistics.median(p['mean'] for p in a),
            'after_median_mean': statistics.median(p['mean'] for p in b),
            'paired_p50_ratios': ratios,
            'paired_mean_ratios': mean_ratios,
            'median_paired_p50_ratio': statistics.median(ratios),
            'bootstrap_95ci_p50_ratio': [lo, hi],
            'processes': processes,
        }
        if metric != 'open_ns':
            record['before_median_open_p50'] = statistics.median(p['open_p50'] for p in a)
            record['after_median_open_p50'] = statistics.median(p['open_p50'] for p in b)
        published = {p['published_bytes'] for p in processes}
        record['published_bytes_identical'] = len(published) == 1
        summary.append(record)
        print(f"{case:38s} {metric:9s} before {record['before_median_p50']/1e6:8.3f} ms  "
              f"after {record['after_median_p50']/1e6:8.3f} ms  ratio {record['median_paired_p50_ratio']:.3f} "
              f"CI [{lo:.3f}, {hi:.3f}]", flush=True)
    with open(os.path.join(OUT, 'summary.json'), 'w') as handle:
        json.dump(summary, handle, indent=1)


if __name__ == '__main__':
    main()
