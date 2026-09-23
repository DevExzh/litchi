#!/usr/bin/env python3
"""Probe measurements for change 0757: the fresh XLS writer's `write_to` alone.

For each case, eight core-pinned processes per case in the order
A B B A A B B A (A = before, B = after), each timing ITERATIONS writes after one
warm-up write, give per-process p50/p95/mean, the median of process p50s per
arm, paired after/before p50 ratios and a percentile bootstrap CI. Each leg is
also run under callgrind at two iteration counts; (Ir(N2) - Ir(N1)) / (N2 - N1)
is the exact instruction count of one write. Allocation counts come from the
probe's counting allocator (identical in every process of a leg).

usage: probe_runs.py OUT_DIR CORE [CASE,... [ROUNDS [pinned]]]

With `pinned`, glibc's trim_threshold and mmap_threshold are pinned at 256 MiB
for the timed processes (not for callgrind), removing heap-trim page faults.
ROUNDS repeats the A B B A group (default 2, eight processes).
"""
import json, os, random, re, statistics, subprocess, sys

OUT, CORE = sys.argv[1], sys.argv[2]
SELECTED = sys.argv[3].split(',') if len(sys.argv) > 3 and sys.argv[3] else None
ROUNDS = int(sys.argv[4]) if len(sys.argv) > 4 else 2
PINNED = len(sys.argv) > 5 and sys.argv[5] == 'pinned'
TUNABLES = 'glibc.malloc.trim_threshold=268435456:glibc.malloc.mmap_threshold=268435456'
SCRATCH = '/home/zhuhe/code/litchi-worktrees/scratch/0757'
PROBE = {'A': f'{SCRATCH}/bin/A/prb', 'B': f'{SCRATCH}/bin/B/prb'}
ORDER_GROUP = ['A', 'B', 'B', 'A']
# (case, timed iterations, callgrind N1, N2)
CASES = [
    ('xls_multi_string/distinct', 30, 1, 3),
    ('xls_multi_string/repeated', 30, 1, 3),
    ('xls_fresh_write_to/tiny', 3000, 20, 120),
    ('xls_fresh_write_to/large', 300, 4, 24),
    ('xls_fresh_write_to/payload-heavy', 40, 1, 4),
    ('doc_fresh_write_to/tiny', 3000, 20, 120),
    ('doc_fresh_write_to/large', 500, 4, 24),
    ('doc_fresh_write_to/payload-heavy', 40, 1, 4),
    ('ppt_fresh_write_to/tiny', 3000, 20, 120),
    ('ppt_fresh_write_to/large', 500, 4, 24),
    ('ppt_fresh_write_to/payload-heavy', 40, 1, 4),
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


def callgrind(leg, case, iterations):
    log = os.path.join(OUT, 'callgrind', f'{case.replace("/", "-")}-{leg}-{iterations}.log')
    result = subprocess.run(['valgrind', '--tool=callgrind', '--callgrind-out-file=/dev/null',
                             PROBE[leg], str(iterations), case],
                            capture_output=True, text=True, check=True)
    with open(log, 'w') as handle:
        handle.write(result.stderr)
    return int(re.search(r'I\s+refs:\s+([\d,]+)', result.stderr).group(1).replace(',', ''))


def main():
    os.makedirs(os.path.join(OUT, 'raw'), exist_ok=True)
    os.makedirs(os.path.join(OUT, 'callgrind'), exist_ok=True)
    summary = []
    order = ORDER_GROUP * ROUNDS
    env = dict(os.environ)
    if PINNED:
        env['GLIBC_TUNABLES'] = TUNABLES
    for case, iterations, n1, n2 in CASES:
        if SELECTED and case not in SELECTED:
            continue
        processes = []
        for index, leg in enumerate(order):
            result = subprocess.run(['taskset', '-c', CORE, PROBE[leg], str(iterations), case],
                                    capture_output=True, text=True, check=True, env=env)
            raw = os.path.join(OUT, 'raw', f'{case.replace("/", "-")}-{index}-{leg}.json')
            with open(raw, 'w') as handle:
                handle.write(result.stdout)
            report = json.loads(result.stdout)
            values = report['ns']
            processes.append({
                'index': index, 'leg': leg, 'raw': os.path.relpath(raw, OUT),
                'p50': pct(values, 0.5), 'p95': pct(values, 0.95), 'mean': statistics.fmean(values),
                'n': len(values), 'output_bytes': report['output_bytes'],
                'output_fnv1a': report['output_fnv1a'],
                'allocations_per_write': report['allocations_per_write'],
                'reallocations_per_write': report['reallocations_per_write'],
                'allocated_bytes_per_write': report['allocated_bytes_per_write'],
                'peak_live_bytes': report['peak_live_bytes'],
            })
        ratios = []
        for i in range(0, len(order), 2):
            first, second = processes[i], processes[i + 1]
            before = first if first['leg'] == 'A' else second
            after = second if second['leg'] == 'B' else first
            ratios.append(after['p50'] / before['p50'])
        lo, hi = bootstrap(ratios)
        a = [p for p in processes if p['leg'] == 'A']
        b = [p for p in processes if p['leg'] == 'B']
        instructions = {}
        for leg in 'AB':
            low, high = callgrind(leg, case, n1), callgrind(leg, case, n2)
            instructions[leg] = (high - low) / (n2 - n1)
        record = {
            'case': case, 'iterations': iterations, 'rounds': ROUNDS, 'pinned_malloc_thresholds': PINNED,
            'before_median_p50_ns': statistics.median(p['p50'] for p in a),
            'after_median_p50_ns': statistics.median(p['p50'] for p in b),
            'before_median_p95_ns': statistics.median(p['p95'] for p in a),
            'after_median_p95_ns': statistics.median(p['p95'] for p in b),
            'paired_p50_ratios': ratios, 'median_paired_p50_ratio': statistics.median(ratios),
            'bootstrap_95ci_p50_ratio': [lo, hi],
            'instructions_per_write': instructions,
            'instructions_ratio': instructions['B'] / instructions['A'],
            'callgrind_iterations': [n1, n2],
            'outputs_before': sorted({(p['output_bytes'], p['output_fnv1a']) for p in a}),
            'outputs_after': sorted({(p['output_bytes'], p['output_fnv1a']) for p in b}),
            'allocations_before': {k: a[0][k] for k in ('allocations_per_write', 'reallocations_per_write', 'allocated_bytes_per_write', 'peak_live_bytes')},
            'allocations_after': {k: b[0][k] for k in ('allocations_per_write', 'reallocations_per_write', 'allocated_bytes_per_write', 'peak_live_bytes')},
            'processes': processes,
        }
        summary.append(record)
        print(f"{case:34s} before {record['before_median_p50_ns']/1e6:9.4f} ms after "
              f"{record['after_median_p50_ns']/1e6:9.4f} ms ratio {record['median_paired_p50_ratio']:.3f} "
              f"CI [{lo:.3f}, {hi:.3f}] Ir {instructions['A']:,.0f} -> {instructions['B']:,.0f} "
              f"({record['instructions_ratio']:.4f}) outputs A {len(record['outputs_before'])} B {len(record['outputs_after'])}",
              flush=True)
    name = 'summary.json' if not SELECTED else f"summary-{'pinned' if PINNED else 'default'}-{ROUNDS}r.json"
    with open(os.path.join(OUT, name), 'w') as handle:
        json.dump(summary, handle, indent=1)


if __name__ == '__main__':
    main()
