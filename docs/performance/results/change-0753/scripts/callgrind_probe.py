#!/usr/bin/env python3
"""Exact instructions per fresh write under callgrind (change 0753).

Runs each leg's allocation probe (the harness's write_fresh_* bodies) for one
case at two iteration counts and reports (Ir(N2) - Ir(N1)) / (N2 - N1), which
excludes process start-up, the probe's warm-up write and argument parsing.

usage: callgrind_probe.py OUT_DIR
"""
import json, os, re, subprocess, sys

OUT = sys.argv[1]
SCRATCH = '/home/zhuhe/code/litchi-worktrees/scratch/0753'
PROBE = {'A': f'{SCRATCH}/bin/A/probe', 'B': f'{SCRATCH}/bin/B/probe'}
CASES = [(f'{kind}_fresh_write_to/{shape}', n1, n2) for kind in ('doc', 'ppt', 'xls')
         for shape, n1, n2 in (('tiny', 20, 120), ('large', 4, 24), ('payload-heavy', 1, 4))]


def instructions(leg, case, iterations):
    log = os.path.join(OUT, f'{case.replace("/", "-")}-{leg}-{iterations}.log')
    subprocess.run(['valgrind', '--tool=callgrind', '--callgrind-out-file=/dev/null',
                    PROBE[leg], str(iterations), case], capture_output=True, text=True, check=True)
    result = subprocess.run(['valgrind', '--tool=callgrind', '--callgrind-out-file=/dev/null',
                             PROBE[leg], str(iterations), case], capture_output=True, text=True, check=True)
    with open(log, 'w') as handle:
        handle.write(result.stderr)
    return int(re.search(r'I\s+refs:\s+([\d,]+)', result.stderr).group(1).replace(',', ''))


def main():
    os.makedirs(OUT, exist_ok=True)
    rows = []
    for case, n1, n2 in CASES:
        row = {'case': case, 'n1': n1, 'n2': n2}
        for leg in 'AB':
            low, high = instructions(leg, case, n1), instructions(leg, case, n2)
            row[leg] = (high - low) / (n2 - n1)
        row['ratio'] = row['B'] / row['A']
        rows.append(row)
        print(f"{case:34s} A {row['A']:15,.0f}  B {row['B']:15,.0f}  B/A {row['ratio']:.4f}", flush=True)
    with open(os.path.join(OUT, 'summary.json'), 'w') as handle:
        json.dump(rows, handle, indent=1)


if __name__ == '__main__':
    main()
