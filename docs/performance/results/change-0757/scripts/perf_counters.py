#!/usr/bin/env python3
"""Per-iteration hardware counters by differencing two process lengths (change 0757).

For each (selector, shape) and leg, runs the harness under `perf stat` with N1
and N2 samples (no warmup) and reports (count(N2) - count(N1)) / (N2 - N1) for
page faults, user/kernel instructions and user/kernel cycles. Order per case:
A B B A, each at both lengths. Per-iteration counts include the harness's
untimed per-iteration work (output comparison), identical in both legs.

usage: perf_counters.py OUT_DIR CORE [selector/shape,...]
"""
import json, os, subprocess, sys

OUT, CORE = sys.argv[1], sys.argv[2]
SCRATCH = '/home/zhuhe/code/litchi-worktrees/scratch/0757'
BIN = {'A': f'{SCRATCH}/bin/A/lpb', 'B': f'{SCRATCH}/bin/B/lpb'}
EVENTS = 'page-faults,instructions:u,cycles:u,instructions:k,cycles:k'
CASES = [
    ('xls_fresh_write_to', 'tiny', 200, 2200), ('xls_fresh_write_to', 'large', 20, 220),
    ('xls_fresh_write_to', 'payload-heavy', 20, 220),
    ('xls_semantic_one_edit_save', 'tiny', 40, 440),
    ('xls_semantic_one_edit_save', 'large', 20, 220),
    ('doc_fresh_write_to', 'tiny', 200, 2200), ('doc_fresh_write_to', 'large', 20, 220),
    ('doc_fresh_write_to', 'payload-heavy', 20, 220),
    ('ppt_fresh_write_to', 'tiny', 200, 2200), ('ppt_fresh_write_to', 'large', 20, 220),
    ('ppt_fresh_write_to', 'payload-heavy', 20, 220),
]


def run(leg, selector, shape, samples, tag):
    stat = os.path.join(OUT, 'raw', f'{selector}-{shape}-{samples}-{tag}-{leg}.txt')
    report = os.path.join(OUT, 'raw', f'{selector}-{shape}-{samples}-{tag}-{leg}.json')
    subprocess.run(['taskset', '-c', CORE, 'perf', 'stat', '-x,', '-e', EVENTS, '-o', stat, BIN[leg],
                    '--case', selector, '--writer-shape', shape, '--samples', str(samples),
                    '--warmup', '0', '--json', report], capture_output=True, text=True, check=True)
    counts = {}
    for line in open(stat):
        parts = line.strip().split(',')
        if len(parts) > 2 and parts[0].replace('.', '').isdigit():
            counts[parts[2]] = float(parts[0])
    return counts


def main():
    os.makedirs(os.path.join(OUT, 'raw'), exist_ok=True)
    summary = []
    selected = sys.argv[3].split(',') if len(sys.argv) > 3 else None
    for selector, shape, short, long in CASES:
        if selected and f'{selector}/{shape}' not in selected:
            continue
        per_leg = {'A': [], 'B': []}
        for tag, leg in enumerate(['A', 'B', 'B', 'A']):
            low = run(leg, selector, shape, short, tag)
            high = run(leg, selector, shape, long, tag)
            per_leg[leg].append({key: (high[key] - low[key]) / (long - short) for key in low})
        row = {'case': selector, 'shape': shape, 'short_samples': short, 'long_samples': long}
        for leg in 'AB':
            row[leg] = {key: sum(r[key] for r in per_leg[leg]) / len(per_leg[leg]) for key in per_leg[leg][0]}
            row[leg + '_runs'] = per_leg[leg]
        summary.append(row)
        a, b = row['A'], row['B']
        print(f"{selector}/{shape}: " + '  '.join(
            f"{key} {a[key]:,.0f} -> {b[key]:,.0f} ({b[key] / a[key] - 1:+.1%})" for key in a), flush=True)
    name = 'summary.json' if not selected else 'summary-' + str(abs(hash(tuple(selected))) % 10000) + '.json'
    with open(os.path.join(OUT, name), 'w') as handle:
        json.dump(summary, handle, indent=1)


if __name__ == '__main__':
    main()
