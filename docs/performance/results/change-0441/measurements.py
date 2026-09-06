#!/usr/bin/env python3
"""Render all normal observations and the repeat-identical allocation vectors."""
import argparse
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SHAPES = {'tiny': 64, 'medium': 4096, 'large': 8192}


def derive():
    summary = json.loads((ROOT / 'summary.json').read_text())
    profiles = json.loads((ROOT / 'profile-summary.json').read_text())
    assert summary['acceptance']['practical_gate_met'] is False
    lines = [
        '# Shared source-projection measurements', '',
        'The original frozen calls/requested-bytes/latency gate failed. The separate',
        '[post-hoc peak-memory review](acceptance-review.md) records the retention',
        'basis. No normal latency, RSS or bounded-memory append improvement is claimed.', '',
        'Each normal row contains 30 samples after three warmups. p50 intervals use',
        '2,000 deterministic bootstrap resamples; p95/p99 use nearest rank, making',
        'p99 the sample maximum. Full means, intervals and raw values remain in summary.json.', '',
        '| Source slides | Phase | p50 ms | p50 95% interval ms | p95 ms | p99 ms |',
        '|---:|---|---:|---|---:|---:|',
    ]
    for row in summary['rows']:
        if row['mode'] != 'normal':
            continue
        stats = row['elapsed_ns']; interval = stats['bootstrap_95']['p50']
        lines.append(f"| {SHAPES[row['shape']]:,} | {row['phase']} | {stats['p50']/1e6:.3f} | {interval['lower']/1e6:.3f}–{interval['upper']/1e6:.3f} | {stats['p95']/1e6:.3f} | {stats['p99']/1e6:.3f} |")
    lines += ['', 'All 60 allocator observations per role and shape agree for the following',
              'counters. Peak is relative to operation entry; retained delta is the endpoint',
              'minus entry, with source and commit still live.', '',
              '| Source slides | Calls before → after | Requested bytes before → after | Peak above entry before → after | Retained delta unchanged |',
              '|---:|---|---|---|---:|']
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
        assert before['retained_live_delta'] == after['retained_live_delta']
        cells = [f"{before[key]:,} → {after[key]:,}" for key in ('allocation_calls', 'allocated_bytes', 'region_peak_above_entry')]
        lines.append(f"| {slides:,} | {' | '.join(cells)} | {before['retained_live_delta']:,} |")
    lines += ['',
              'Medium/large peak falls 9.834%/9.857%. Allocation calls fall only',
              '1.398%/1.403%, and requested bytes 1.708%/1.778%. Normal medium/large p50',
              'is 0.664–1.279% slower. Requested bytes do not measure physical memory copies.', '',
              f"Matched adverse >5% flags: {len(summary['review_flags'])}. Repeat flags: {len(summary['repeat_flags'])}."]
    for row in summary['repeat_flags']:
        lines.append(f"The {row['role']} {row['mode']} {row['shape']} {row['metric']} repeat changes {row['relative_percent']:+.3f}%.")
    for role in ('before', 'after'):
        rss = [row['process_memory']['gnu_time_peak_rss_bytes'] for row in summary['rows'] if row['role'] == role]
        lines.append(f"Whole-process {role} peak RSS spans {min(rss):,}–{max(rss):,} bytes.")
    lines += ['', 'Four whole-process profiles include setup, warmups and Rust oracle work.']
    for name, change in profiles['counter_change_percent'].items():
        if change is not None:
            lines.append(f"`{name}` changes {change:+.3f}%.")
    lines.append('Zero L1 values support no cache-miss interpretation.')
    for role, row in profiles['roles'].items():
        lines.append(f"{role.capitalize()} symbolization retains {row['perf_report_addr2line_warnings']} report and {row['perf_script_addr2line_warnings']} script addr2line warnings.")
    lines += ['', 'See [summary.json](summary.json), [profile-summary.json](profile-summary.json),',
              'and [decision.json](decision.json). These observations cover the named owned',
              'append fixtures on this machine and build; the broader non-iWork goal remains open.', '']
    return '\n'.join(lines)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    text = derive()
    path = ROOT / 'measurements.md'
    if args.check:
        assert path.read_text() == text
    else:
        path.write_text(text)
    print('VALID: measurement tables and constant allocation vectors')
