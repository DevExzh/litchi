#!/usr/bin/env python3
"""Heap-state check for change 0757: one case under glibc's default malloc
thresholds and with trim_threshold and mmap_threshold pinned at 256 MiB, in the
order A B B A for each setting. Reports per-iteration page faults and user
instructions (differenced 20 vs 220 samples) and the timed p50 of the 220-sample
process.

usage: tunables_check.py OUT_DIR CORE CASE SHAPE
"""
import json, os, statistics, subprocess, sys

OUT, CORE, CASE, SHAPE = sys.argv[1:5]
SCRATCH = '/home/zhuhe/code/litchi-worktrees/scratch/0757'
BIN = {'A': f'{SCRATCH}/bin/A/lpb', 'B': f'{SCRATCH}/bin/B/lpb'}
PINNED = 'glibc.malloc.trim_threshold=268435456:glibc.malloc.mmap_threshold=268435456'


def run(leg, samples, tag, pinned):
    env = dict(os.environ)
    if pinned:
        env['GLIBC_TUNABLES'] = PINNED
    base = os.path.join(OUT, f'{CASE}-{SHAPE}-{"pinned" if pinned else "default"}-{tag}-{leg}-{samples}')
    subprocess.run(['taskset', '-c', CORE, 'perf', 'stat', '-x,', '-e', 'page-faults,instructions:u,cycles:u',
                    '-o', base + '.txt', BIN[leg], '--case', CASE, '--writer-shape', SHAPE,
                    '--samples', str(samples), '--warmup', '0', '--json', base + '.json'],
                   env=env, capture_output=True, text=True, check=True)
    counts = {}
    for line in open(base + '.txt'):
        parts = line.strip().split(',')
        if len(parts) > 2 and parts[0].replace('.', '').isdigit():
            counts[parts[2]] = float(parts[0])
    report = json.load(open(base + '.json'))
    samples_ns = [r for r in report['results'] if r['case'] == CASE][0]['elapsed_ns']['samples']
    return counts, statistics.median(samples_ns)


rows = []
for pinned in (False, True):
    for tag, leg in enumerate('ABBA'):
        low, _ = run(leg, 20, tag, pinned)
        high, p50 = run(leg, 220, tag, pinned)
        row = {'setting': 'pinned' if pinned else 'default', 'leg': leg, 'p50_ms': p50 / 1e6}
        row.update({key: (high[key] - low[key]) / 200 for key in low})
        rows.append(row)
        print(json.dumps(row), flush=True)
json.dump(rows, open(os.path.join(OUT, f'summary-{CASE}-{SHAPE}.json'), 'w'), indent=1)
