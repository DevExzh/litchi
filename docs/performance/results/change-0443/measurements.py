#!/usr/bin/env python3
"""Render matched latency, allocation, uncertainty and retained review flags."""
import argparse
import json
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SHAPES = {'tiny': 64, 'medium': 4096, 'large': 8192}


def derive():
    summary = json.loads((ROOT / 'summary.json').read_text())
    profiles = json.loads((ROOT / 'profile-summary.json').read_text())
    lines = [
        '# Compact fragment scanner measurements', '',
        f"Frozen acceptance: normal-p50 gate={summary['acceptance']['normal_p50_gate']}; allocation/peak gate={summary['acceptance']['allocator_requested_call_or_peak_gate']}; practical gate={summary['acceptance']['practical_gate_met']}.",
        'The matrix retains 24 reports and 720 samples in A1/B1/B2/A2 order.',
        'Each normal row has 30 samples after three warmups. p50 intervals use 2,000',
        'deterministic bootstrap resamples; p95/p99 use nearest rank, making p99 the',
        'sample maximum. These intervals do not eliminate temporal grouping effects.', '',
        '| Source slides | Phase | p50 ms | p50 95% interval ms | p95 ms | p99 ms |',
        '|---:|---|---:|---|---:|---:|',
    ]
    for row in summary['rows']:
        if row['mode'] != 'normal':
            continue
        stats = row['elapsed_ns']; interval = stats['bootstrap_95']['p50']
        lines.append(f"| {SHAPES[row['shape']]:,} | {row['phase']} | {stats['p50']/1e6:.3f} | {interval['lower']/1e6:.3f}–{interval['upper']/1e6:.3f} | {stats['p95']/1e6:.3f} | {stats['p99']/1e6:.3f} |")
    pairs = [row for row in summary['comparisons'] if row['mode'] == 'normal' and row['metric'] == 'elapsed_ns.p50']
    lines.append('')
    for shape in SHAPES:
        changes = [row for row in pairs if row['shape'] == shape]
        lines.append(f"{shape.capitalize()} p50 changes R1/R2: " + ' / '.join(f"{row['relative_percent']:+.3f}%" for row in changes) + '.')
    ratio = math.exp(sum(math.log(row['before']/row['after']) for row in pairs)/len(pairs))
    lines += [f'The geometric mean of the six explicitly paired p50 before/after ratios is {ratio:.4f}×.',
              'Each ratio normalizes one shape and repeat to its own baseline; this is not',
              'a geometric mean across different Office workflows.', '',
              'All 60 allocator observations per role and shape agree for each counter below.',
              'Source and committed snapshot remain live at the endpoint;',
              'append still materializes the document.', '',
              '| Source slides | Calls before → after | Requested bytes before → after | Peak above entry before → after | Retained delta before → after |',
              '|---:|---|---|---:|---:|']
    for shape, slides in SHAPES.items():
        roles = {}
        for role in ('before', 'after'):
            roles[role] = {}
            rows = [row for row in summary['rows'] if row['shape'] == shape and row['role'] == role and row['mode'] == 'allocator']
            for key in ('allocation_calls', 'allocated_bytes', 'region_peak_above_entry', 'retained_live_delta'):
                values = [value for row in rows for value in row['allocation']['chronological_vectors'][key]]
                assert len(values) == 60 and len(set(values)) == 1
                roles[role][key] = values[0]
        before, after = roles['before'], roles['after']
        cells = [f"{before[key]:,} → {after[key]:,}" for key in ('allocation_calls', 'allocated_bytes', 'region_peak_above_entry', 'retained_live_delta')]
        lines.append(f"| {slides:,} | {' | '.join(cells)} |")
    lines += ['', 'Requested bytes are allocation traffic, not physical-copy counts.',
              'No bounded append or general RSS improvement is claimed.', '',
              f"Matched adverse >5% flags: {len(summary['review_flags'])}. Repeat flags: {len(summary['repeat_flags'])}."]
    for row in summary['review_flags']:
        lines.append(f"Matched {row['mode']} {row['shape']} {row['repeat']} {row['metric']} changes {row['relative_percent']:+.3f}%.")
    for row in summary['repeat_flags']:
        lines.append(f"The {row['role']} {row['mode']} {row['shape']} {row['metric']} repeat changes {row['relative_percent']:+.3f}%.")
    tail = next(row for row in summary['comparisons'] if row['mode'] == 'normal' and row['shape'] == 'large' and row['repeat'] == 'R2' and row['metric'] == 'elapsed_ns.p99')
    lines += [f"The large R2 paired p99 comparison is {tail['before']/1e6:.3f} → {tail['after']/1e6:.3f} ms ({tail['relative_percent']:+.3f}%).",
              'The retained repeat variation and 30-sample tail resolution preclude a general',
              'tail-latency improvement claim. Instrumented timing is not used to establish a normal speedup.']
    for role in ('before', 'after'):
        rss = [row['process_memory']['gnu_time_peak_rss_bytes'] for row in summary['rows'] if row['role'] == role]
        lines.append(f"Whole-process {role} peak RSS spans {min(rss):,}–{max(rss):,} bytes.")
    lines += ['', 'Four whole-process profiles include setup, warmups and Rust oracle work.']
    for name, change in profiles['counter_change_percent'].items():
        if change is not None:
            lines.append(f"`{name}` changes {change:+.3f}%.")
    lines.append('Zero L1 values support no cache-miss interpretation; these are not operation-only causal fractions.')
    for role, row in profiles['roles'].items():
        lines.append(f"{role.capitalize()} symbolization retains {row['perf_report_addr2line_warnings']} report and {row['perf_script_addr2line_warnings']} script addr2line warnings.")
    lines += ['', 'See [summary.json](summary.json), [profile-summary.json](profile-summary.json),',
              'and [decision.json](decision.json). The broader non-iWork goal remains active.', '']
    return '\n'.join(lines)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(); parser.add_argument('--check', action='store_true'); args = parser.parse_args()
    text = derive(); path = ROOT / 'measurements.md'
    if args.check:
        assert path.read_text() == text
    else:
        path.write_text(text)
    print('VALID: measurement tables and allocation vectors')
