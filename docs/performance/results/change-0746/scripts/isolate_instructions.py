#!/usr/bin/env python3
"""Per-operation user-space instruction and cycle counts by isolation pairs.

For each (fixture, operation, leg) the probe runs twice under `perf stat`,
once with SMALL and once with LARGE timed iterations (no warmups), and the
difference divided by LARGE-SMALL prices one iteration with every untimed
setup (file read, target discovery, process start) cancelled. Commit
operations use --reuse-source so an iteration is the staged edit, the commit
and the publication alone; open operations time one complete open each.

usage: isolate_instructions.py OUT_DIR CORE
"""
import json, os, subprocess, sys

OUT = sys.argv[1]
CORE = sys.argv[2]
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BIN = {'before': os.path.join(ROOT, 'bin', 'probe-before'), 'after': os.path.join(ROOT, 'bin', 'probe-after')}
FIX = os.path.join(ROOT, 'fixtures')
SMALL, LARGE = 4, 24
CASES = [
    ('54016.xls', 'open'), ('54016.xls', 'number-plan'), ('54016.xls', 'number-source-backed'),
    ('54016.xls', 'number-generic'), ('54016.xls', 'string-generic'), ('54016.xls', 'comments-open'),
    ('54016.xls', 'visibility-open'), ('54016.xls', 'reader-open'),
    ('xls-large.xls', 'open'), ('xls-large.xls', 'number-plan'), ('xls-large.xls', 'number-source-backed'),
    ('xls-large.xls', 'number-generic'), ('xls-large.xls', 'comments-open'),
    ('xls-large.xls', 'visibility-open'), ('xls-large.xls', 'reader-open'),
]
COMMITS = {'number-plan', 'number-source-backed', 'number-generic', 'string-generic'}


def counts(leg, fixture, operation, samples):
    command = ['perf', 'stat', '-x', ',', '-e', 'instructions:u,cycles:u', 'taskset', '-c', CORE,
               BIN[leg], '--input', os.path.join(FIX, fixture), '--operation', operation,
               '--warmups', '0', '--samples', str(samples)]
    if operation in COMMITS:
        command.append('--reuse-source')
    result = subprocess.run(command, capture_output=True, text=True, check=True)
    values = {}
    for line in result.stderr.splitlines():
        parts = line.split(',')
        if len(parts) > 2 and parts[2].startswith(('instructions', 'cycles')):
            values[parts[2].split(':')[0]] = int(parts[0])
    return values


def main():
    os.makedirs(OUT, exist_ok=True)
    rows = []
    for fixture, operation in CASES:
        row = {'fixture': fixture, 'operation': operation, 'small': SMALL, 'large': LARGE}
        for leg in ('before', 'after'):
            small = counts(leg, fixture, operation, SMALL)
            large = counts(leg, fixture, operation, LARGE)
            row[leg] = {
                'raw_small': small, 'raw_large': large,
                'instructions_per_op': (large['instructions'] - small['instructions']) / (LARGE - SMALL),
                'cycles_per_op': (large['cycles'] - small['cycles']) / (LARGE - SMALL),
            }
        for metric in ('instructions_per_op', 'cycles_per_op'):
            row[metric + '_change'] = row['after'][metric] / row['before'][metric] - 1
        rows.append(row)
        print(f"{fixture:14s} {operation:21s} instr {row['before']['instructions_per_op']:14,.0f} -> "
              f"{row['after']['instructions_per_op']:14,.0f} ({row['instructions_per_op_change']*100:+6.1f}%)  "
              f"cycles {row['before']['cycles_per_op']:14,.0f} -> {row['after']['cycles_per_op']:14,.0f} "
              f"({row['cycles_per_op_change']*100:+6.1f}%)", flush=True)
    with open(os.path.join(OUT, 'isolation.json'), 'w') as handle:
        json.dump(rows, handle, indent=1)


if __name__ == '__main__':
    main()
