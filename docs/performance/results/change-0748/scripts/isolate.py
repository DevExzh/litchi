#!/usr/bin/env python3
"""Per-iteration user-space instructions and cycles for change 0748.

For each case and leg, runs the same binary with a short and a long iteration
count under ``perf stat -e instructions:u,cycles:u`` (pinned with taskset) and
reports ``(long - short) / (long_n - short_n)``: the work of one timed
iteration, free of process startup and untimed corpus construction. Harness
per-sample counts still include any untimed per-sample work the selector does
(for example cloning its corpus). Each (leg, length) is repeated REPEATS times
in A B B A order and the median is kept.

usage: isolate.py OUT_JSON CORE [case,case,...]
"""
import json
import os
import statistics
import subprocess
import sys

OUT = sys.argv[1]
CORE = sys.argv[2]
SCRATCH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BIN = {
    leg: {
        'harness': os.path.join(SCRATCH, 'bin', name, 'litchi-perf-baseline'),
        'probe': os.path.join(SCRATCH, 'bin', name, 'xls_edit_probe'),
    }
    for leg, name in (('A', 'before'), ('B', 'after'))
}
XLS_54016 = os.path.join(SCRATCH, 'fixtures', '54016.xls')
XLS_LARGE = os.path.join(SCRATCH, 'fixtures', 'xls-large.xls')
SHORT = int(os.environ.get('ISOLATE_SHORT', '4'))
LONG = int(os.environ.get('ISOLATE_LONG', '24'))
REPEATS = int(os.environ.get('ISOLATE_REPEATS', '3'))
LARGE = ['--writer-shape', 'large']
FS = ['--filesystem-cache', 'warm', '--filesystem-root', os.path.join(SCRATCH, 'fsroot')]

CASES = [
    ('probe/54016/number-generic', 'probe', ['--input', XLS_54016, '--operation', 'number-generic', '--reuse-source']),
    ('probe/54016/number-source-backed', 'probe', ['--input', XLS_54016, '--operation', 'number-source-backed', '--reuse-source']),
    ('probe/54016/number-plan', 'probe', ['--input', XLS_54016, '--operation', 'number-plan', '--reuse-source']),
    ('probe/54016/open', 'probe', ['--input', XLS_54016, '--operation', 'open']),
    ('probe/xls-large/number-generic', 'probe', ['--input', XLS_LARGE, '--operation', 'number-generic', '--reuse-source']),
    ('xls_semantic_one_edit_save/large', 'harness', ['--case', 'xls_semantic_one_edit_save', *LARGE]),
    ('xls_visibility_eager_edit_save', 'harness', ['--case', 'xls_visibility_eager_edit_save']),
    ('xls_visibility_source_backed_edit_save', 'harness', ['--case', 'xls_visibility_source_backed_edit_save']),
    ('xls_numeric_eager_rk_mulrk_edit_save', 'harness', ['--case', 'xls_numeric_eager_rk_mulrk_edit_save']),
    ('xls_numeric_plan_only_rk_mulrk_edit_save', 'harness', ['--case', 'xls_numeric_plan_only_rk_mulrk_edit_save']),
    ('xls_numeric_source_backed_rk_mulrk_edit_save', 'harness', ['--case', 'xls_numeric_source_backed_rk_mulrk_edit_save']),
    ('xls_comments_source_backed_edit_save', 'harness', ['--case', 'xls_comments_source_backed_edit_save']),
    ('xls_comments_eager_edit_save', 'harness', ['--case', 'xls_comments_eager_edit_save']),
    ('ppt_source_backed_one_shape_text/large', 'harness', ['--case', 'ppt_source_backed_one_shape_text', *LARGE]),
    ('doc_semantic_one_edit_save/large', 'harness', ['--case', 'doc_semantic_one_edit_save', *LARGE]),
    ('cfb_file_same_length_overlay_atomic_save', 'harness', ['--case', 'cfb_file_same_length_overlay_atomic_save', *FS]),
    ('cfb_file_owned_same_length_overlay_atomic_save', 'harness', ['--case', 'cfb_file_owned_same_length_overlay_atomic_save', *FS]),
]


def counters(leg, kind, arguments, iterations):
    if kind == 'probe':
        tail = [BIN[leg]['probe'], *arguments, '--warmups', '0', '--samples', str(iterations)]
    else:
        tail = [BIN[leg]['harness'], *arguments, '--warmup', '0', '--samples', str(iterations),
                '--json', os.devnull]
    command = ['perf', 'stat', '-x', ',', '-e', 'instructions:u,cycles:u', '--',
               'taskset', '-c', CORE, *tail]
    completed = subprocess.run(command, capture_output=True, text=True, check=True)
    values = {}
    for line in completed.stderr.splitlines():
        fields = line.split(',')
        if len(fields) > 2 and fields[2].startswith(('instructions', 'cycles')):
            values[fields[2].split(':')[0]] = int(fields[0])
    return values['instructions'], values['cycles']


def main():
    selected = sys.argv[3].split(',') if len(sys.argv) > 3 else None
    results = []
    for label, kind, arguments in CASES:
        if selected and label not in selected:
            continue
        per_leg = {}
        runs = {('A', SHORT): [], ('A', LONG): [], ('B', SHORT): [], ('B', LONG): []}
        for _ in range(REPEATS):
            for leg in ('A', 'B', 'B', 'A'):
                for iterations in (SHORT, LONG):
                    runs[(leg, iterations)].append(counters(leg, kind, arguments, iterations))
        for leg in ('A', 'B'):
            short_i = statistics.median(value[0] for value in runs[(leg, SHORT)])
            long_i = statistics.median(value[0] for value in runs[(leg, LONG)])
            short_c = statistics.median(value[1] for value in runs[(leg, SHORT)])
            long_c = statistics.median(value[1] for value in runs[(leg, LONG)])
            per_leg[leg] = {
                'instructions_per_iteration': (long_i - short_i) / (LONG - SHORT),
                'cycles_per_iteration': (long_c - short_c) / (LONG - SHORT),
                'raw': {f'{n}': [list(v) for v in runs[(leg, n)]] for n in (SHORT, LONG)},
            }
        a, b = per_leg['A'], per_leg['B']
        record = {
            'case': label,
            'before_instructions': a['instructions_per_iteration'],
            'after_instructions': b['instructions_per_iteration'],
            'instructions_ratio': b['instructions_per_iteration'] / a['instructions_per_iteration'],
            'before_cycles': a['cycles_per_iteration'],
            'after_cycles': b['cycles_per_iteration'],
            'cycles_ratio': b['cycles_per_iteration'] / a['cycles_per_iteration'],
            'legs': per_leg,
        }
        results.append(record)
        print(f"{label:48s} instr {a['instructions_per_iteration'] / 1e6:9.3f}M -> "
              f"{b['instructions_per_iteration'] / 1e6:9.3f}M ({record['instructions_ratio']:.3f})  "
              f"cycles {a['cycles_per_iteration'] / 1e6:9.3f}M -> {b['cycles_per_iteration'] / 1e6:9.3f}M "
              f"({record['cycles_ratio']:.3f})", flush=True)
    with open(OUT, 'w') as handle:
        json.dump({'core': CORE, 'short': SHORT, 'long': LONG, 'repeats': REPEATS,
                   'results': results}, handle, indent=1)


if __name__ == '__main__':
    main()
