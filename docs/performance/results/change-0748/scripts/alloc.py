#!/usr/bin/env python3
"""Deterministic allocation counts for change 0748.

Runs the change-0746 probe built with `--features alloc-count` for each leg.
The probe counts every allocation of the first measured iteration (open,
stage, commit and publish for the edit operations) and reports calls, bytes
and peak live bytes. Two runs per leg confirm determinism.

usage: alloc.py OUT_JSON CORE
"""
import json
import os
import subprocess
import sys

OUT = sys.argv[1]
CORE = sys.argv[2]
SCRATCH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BIN = {
    'before': os.path.join(SCRATCH, 'bin', 'before', 'xls_edit_probe_alloc'),
    'after': os.path.join(SCRATCH, 'bin', 'after', 'xls_edit_probe_alloc'),
}
FIXTURES = {
    '54016': os.path.join(SCRATCH, 'fixtures', '54016.xls'),
    'xls-large': os.path.join(SCRATCH, 'fixtures', 'xls-large.xls'),
}
OPERATIONS = ['number-generic', 'string-generic', 'number-source-backed', 'number-plan', 'open']
FIELDS = ['allocations', 'allocated_bytes', 'peak_live_bytes', 'open_allocations',
          'open_allocated_bytes', 'open_peak_live_bytes']


def main():
    results = []
    for fixture, path in FIXTURES.items():
        for operation in OPERATIONS:
            if fixture == 'xls-large' and operation == 'string-generic':
                continue
            row = {'fixture': fixture, 'operation': operation}
            for leg, binary in BIN.items():
                runs = []
                for _ in range(2):
                    completed = subprocess.run(
                        ['taskset', '-c', CORE, binary, '--input', path, '--operation', operation,
                         '--warmups', '1', '--samples', '1'],
                        capture_output=True, text=True, check=True)
                    report = json.loads(completed.stdout)
                    runs.append({field: report.get(field) for field in FIELDS})
                row[leg] = runs[0]
                row[f'{leg}_deterministic'] = runs[0] == runs[1]
            results.append(row)
            b, a = row['before'], row['after']
            print(f"{fixture:9s} {operation:22s} calls {b['allocations']:>7} -> {a['allocations']:>7}  "
                  f"bytes {b['allocated_bytes']:>11} -> {a['allocated_bytes']:>11}  "
                  f"peak {b['peak_live_bytes']:>10} -> {a['peak_live_bytes']:>10}  "
                  f"det {row['before_deterministic']}/{row['after_deterministic']}", flush=True)
    with open(OUT, 'w') as handle:
        json.dump(results, handle, indent=1)


if __name__ == '__main__':
    main()
