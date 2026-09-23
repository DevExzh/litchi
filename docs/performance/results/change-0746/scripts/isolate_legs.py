#!/usr/bin/env python3
"""Isolation-pair instructions and cycles per operation for several probe legs.

usage: isolate_legs.py OUT_JSON CORE LEG=BIN[,LEG=BIN...] CASE[;CASE...]
CASE = fixture:operation:small:large ; commit operations add --reuse-source.
"""
import json, os, subprocess, sys
OUT, CORE, LEGS, SPEC = sys.argv[1:5]
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIX = os.path.join(ROOT, 'fixtures')
COMMITS = {'number-plan', 'number-source-backed', 'number-generic', 'string-generic'}
legs = [leg.split('=', 1) for leg in LEGS.split(',')]

def counts(binary, fixture, operation, samples):
    command = ['perf', 'stat', '-x', ',', '-e', 'instructions:u,cycles:u', 'taskset', '-c', CORE, binary,
               '--input', os.path.join(FIX, fixture), '--operation', operation, '--warmups', '0',
               '--samples', str(samples)]
    if operation in COMMITS:
        command.append('--reuse-source')
    result = subprocess.run(command, capture_output=True, text=True, check=True)
    values = {}
    for line in result.stderr.splitlines():
        parts = line.split(',')
        if len(parts) > 2 and parts[2].startswith(('instructions', 'cycles')):
            values[parts[2].split(':')[0]] = int(parts[0])
    return values

rows = []
for case in SPEC.split(';'):
    fixture, operation, small, large = case.split(':')
    small, large = int(small), int(large)
    row = {'fixture': fixture, 'operation': operation, 'small': small, 'large': large}
    line = f"{fixture:26s} {operation:21s}"
    for name, binary in legs:
        a, b = counts(binary, fixture, operation, small), counts(binary, fixture, operation, large)
        row[name] = {'raw_small': a, 'raw_large': b,
                     'instructions_per_op': (b['instructions'] - a['instructions']) / (large - small),
                     'cycles_per_op': (b['cycles'] - a['cycles']) / (large - small)}
        line += f" | {name} {row[name]['instructions_per_op']/1e6:8.2f} Mi {row[name]['cycles_per_op']/1e6:7.2f} Mc"
    rows.append(row)
    print(line, flush=True)
with open(OUT, 'w') as handle:
    json.dump(rows, handle, indent=1)
