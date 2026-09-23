#!/usr/bin/env python3
"""ABBA-interleaved, core-pinned probe timing between two named probe legs.

usage: abba_legs.py OUT_DIR CORE A_BIN B_BIN CASE[;CASE...]
CASE = fixture:operation:warmups:samples:metric

Each case runs 8 processes, A B B A A B B A. The summary reports per-process
p50/p95/mean, the median of process p50s per leg, the four adjacent-pair B/A
p50 ratios (and mean ratios) and their minimum and maximum. With four pairs a
percentile bootstrap interval over the ratios is their min-max range, so the
range is reported directly.
"""
import json, os, statistics, subprocess, sys

OUT, CORE, A_BIN, B_BIN, SPEC = sys.argv[1:6]
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIX = os.path.join(ROOT, 'fixtures')
ORDER = ['A', 'B', 'B', 'A', 'A', 'B', 'B', 'A']
BIN = {'A': A_BIN, 'B': B_BIN}


def pct(values, q):
    values = sorted(values)
    k = (len(values) - 1) * q
    lo, hi = int(k), min(int(k) + 1, len(values) - 1)
    return values[lo] + (values[hi] - values[lo]) * (k - lo)


def main():
    os.makedirs(os.path.join(OUT, 'raw'), exist_ok=True)
    summary = []
    for case in SPEC.split(';'):
        fixture, operation, warmups, samples, metric = case.split(':')
        processes = []
        for index, leg in enumerate(ORDER):
            raw = os.path.join(OUT, 'raw', f'{fixture}-{operation}-{index}-{leg}.json')
            command = ['taskset', '-c', CORE, BIN[leg], '--input', os.path.join(FIX, fixture),
                       '--operation', operation, '--warmups', warmups, '--samples', samples]
            result = subprocess.run(command, capture_output=True, text=True, check=True)
            with open(raw, 'w') as handle:
                handle.write(result.stdout)
            values = json.loads(result.stdout)[metric]
            processes.append({'index': index, 'leg': leg, 'raw': os.path.relpath(raw, OUT),
                              'p50': pct(values, 0.5), 'p95': pct(values, 0.95),
                              'mean': statistics.fmean(values), 'n': len(values)})
        a = [p for p in processes if p['leg'] == 'A']
        b = [p for p in processes if p['leg'] == 'B']
        ratios, mean_ratios = [], []
        for i in range(0, len(ORDER), 2):
            first, second = processes[i], processes[i + 1]
            before = first if first['leg'] == 'A' else second
            after = second if second['leg'] == 'B' else first
            ratios.append(after['p50'] / before['p50'])
            mean_ratios.append(after['mean'] / before['mean'])
        record = {
            'case': f'{fixture}:{operation}', 'metric': metric, 'warmups': int(warmups),
            'samples': int(samples), 'a_bin': A_BIN, 'b_bin': B_BIN,
            'a_median_p50': statistics.median(p['p50'] for p in a),
            'b_median_p50': statistics.median(p['p50'] for p in b),
            'a_median_p95': statistics.median(p['p95'] for p in a),
            'b_median_p95': statistics.median(p['p95'] for p in b),
            'a_median_mean': statistics.median(p['mean'] for p in a),
            'b_median_mean': statistics.median(p['mean'] for p in b),
            'paired_p50_ratios': ratios, 'paired_mean_ratios': mean_ratios,
            'median_paired_p50_ratio': statistics.median(ratios),
            'paired_p50_ratio_range': [min(ratios), max(ratios)],
            'processes': processes,
        }
        summary.append(record)
        print(f"{record['case']:40s} A {record['a_median_p50']/1e6:9.3f} ms  B {record['b_median_p50']/1e6:9.3f} ms  "
              f"ratio {record['median_paired_p50_ratio']:.3f} range [{min(ratios):.3f}, {max(ratios):.3f}]",
              flush=True)
    with open(os.path.join(OUT, 'summary.json'), 'w') as handle:
        json.dump(summary, handle, indent=1)


if __name__ == '__main__':
    main()
