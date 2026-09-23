#!/usr/bin/env python3
"""Per-sample user-space instructions/cycles of harness selectors by isolation
pairs: the same selector with SMALL and LARGE samples (no warmup) under
`perf stat`, differenced and divided by LARGE-SMALL, cancels process start and
corpus construction.

usage: isolate_harness.py OUT_DIR CORE SELECTOR[,SELECTOR...]
"""
import json, os, subprocess, sys
OUT, CORE, SELECTORS = sys.argv[1], sys.argv[2], sys.argv[3].split(',')
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BIN = {'before': os.environ.get('HARNESS_BEFORE', os.path.join(ROOT, 'bin', 'harness-before')), 'after': os.path.join(ROOT, 'bin', 'harness-after')}
EXTRA = {'xls_semantic_one_edit_save': ['--writer-shape', 'large'], 'xls_semantic_noop_edit_save': ['--writer-shape', 'large'],
         'xls_semantic_open': ['--writer-shape', 'large']}
SMALL, LARGE = 4, 24

def counts(leg, selector, samples):
    json_path = os.path.join(OUT, f'{selector}-{leg}-{samples}.json')
    command = ['perf', 'stat', '-x', ',', '-e', 'instructions:u,cycles:u', 'taskset', '-c', CORE, BIN[leg],
               '--case', selector, '--samples', str(samples), '--warmup', '0', '--json', json_path,
               *EXTRA.get(selector, [])]
    result = subprocess.run(command, capture_output=True, text=True, check=True)
    values = {}
    for line in result.stderr.splitlines():
        parts = line.split(',')
        if len(parts) > 2 and parts[2].startswith(('instructions', 'cycles')):
            values[parts[2].split(':')[0]] = int(parts[0])
    return values

os.makedirs(OUT, exist_ok=True)
rows = []
for selector in SELECTORS:
    row = {'selector': selector, 'small': SMALL, 'large': LARGE}
    for leg in ('before', 'after'):
        small, large = counts(leg, selector, SMALL), counts(leg, selector, LARGE)
        row[leg] = {'raw_small': small, 'raw_large': large,
                    'instructions_per_sample': (large['instructions'] - small['instructions']) / (LARGE - SMALL),
                    'cycles_per_sample': (large['cycles'] - small['cycles']) / (LARGE - SMALL)}
    for metric in ('instructions_per_sample', 'cycles_per_sample'):
        row[metric + '_change'] = row['after'][metric] / row['before'][metric] - 1
    rows.append(row)
    print(f"{selector:44s} instr {row['before']['instructions_per_sample']:14,.0f} -> {row['after']['instructions_per_sample']:14,.0f} "
          f"({row['instructions_per_sample_change']*100:+6.2f}%)  cycles {row['before']['cycles_per_sample']:14,.0f} -> "
          f"{row['after']['cycles_per_sample']:14,.0f} ({row['cycles_per_sample_change']*100:+6.2f}%)", flush=True)
with open(os.path.join(OUT, 'isolation-harness.json'), 'w') as handle:
    json.dump(rows, handle, indent=1)
