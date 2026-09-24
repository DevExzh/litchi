#!/usr/bin/env python3
"""Derive the date/time report from the validated capture; never runs a benchmark."""
from collections import defaultdict
import hashlib
import importlib.util
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
RESULTS = HERE / 'performance/results'
FIELDS = ('alloc_calls', 'requested_bytes', 'released_bytes', 'live_before', 'live_after',
          'peak_live_delta', 'work', 'memory_retained', 'reference_reads', 'output_bytes')


def load(path):
    return json.loads(path.read_text())


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path, value):
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')


def main():
    subprocess.run([sys.executable, str(HERE / "performance/analyze_profile.py")], check=True)
    spec = importlib.util.spec_from_file_location('profile_analysis', HERE / 'performance/analyze_profile.py')
    analyzer = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(analyzer)
    analysis = load(RESULTS / 'profile-analysis.json')
    summary = load(RESULTS / 'capture-summary.json')
    assert analysis['status'] == 'ok'
    captures = {}
    summaries = {}
    for label in ('baseline-' + summary['baseline_commit'], 'candidate-final'):
        groups = defaultdict(list)
        for line in (RESULTS / label / 'measurements.jsonl').read_text().splitlines():
            row = json.loads(line)
            groups[row['case'] + '/' + row['phase']].append(row)
        captures[label] = groups
        summaries[label] = []
        for key, rows in sorted(groups.items()):
            normalized = analyzer.group_summary([
                {'row': row, 'sample': row['samples'][0], 'rss_kib': row['rss_kib']} for row in rows
            ])
            normalized.update(case=rows[0]['case'], phase=rows[0]['phase'], operation=rows[0]['operation'])
            summaries[label].append(normalized)
    baseline = captures['baseline-' + summary['baseline_commit']]
    candidate = captures['candidate-final']
    differences = []
    for group, rows in baseline.items():
        for field in FIELDS:
            before = sorted({row['samples'][0][field] for row in rows})
            after = sorted({row['samples'][0][field] for row in candidate[group]})
            if before != after:
                differences.append({'group': group, 'field': field, 'baseline': before, 'candidate': after})
    assert not differences, differences
    flags = analysis['comparison']['flags']
    median = [flag for flag in flags if flag['quantile'] == 'median']
    notes = [
        'Accepted with disclosed threshold flags for this function-support batch; no overall speedup or causal regression claim.',
        'Latency includes cached expected-value comparisons, complete checksumming and result drop; independent oracle generation is outside timing.',
        'Baseline precedes candidate; CPU affinity is not pinned and system load is recorded without a rejection threshold.',
        'Bootstrap intervals are descriptive independent resamples of fifteen child samples, not proof of a causal effect.',
        'RSS is a process observation; equal allocation counters do not establish a cause for RSS differences.',
        'Hardware-counter and concurrency scaling improvements are not claimed by this synchronous evaluator capture.',
    ]
    report = {'schema': 'ods-formula-date-time-performance-report-v1', 'status': 'ok',
              'contract_sha256': summary['contract_sha256'],
              'freeze_sha256': summary['candidate_freeze']['sha256'],
              'analysis_sha256': digest(RESULTS / 'profile-analysis.json'),
              'captures': summary['captures'], 'groups': summaries,
              'comparison': analysis['comparison'], 'disposition': notes,
              'accounting_sets_unchanged': True}
    write(RESULTS / 'performance-report.json', report)
    audit = {'status': 'verified', 'rows': analysis['baseline_rows'] + analysis['candidate_rows'],
             'matched_groups': len(baseline), 'accounting_fields': FIELDS,
             'accounting_sets_unchanged': True, 'differences': differences,
             'report_sha256': digest(RESULTS / 'performance-report.json'),
             'analysis_sha256': digest(RESULTS / 'profile-analysis.json'),
             'summarizer_sha256': digest(Path(__file__)), 'median_flags': median,
             'disposition': notes}
    write(HERE / 'root-performance-audit.json', audit)
    lines = ['# Date/time performance evidence', '',
             f"Validated {audit['rows']:,} samples: {analysis['baseline_rows']:,} baseline and {analysis['candidate_rows']:,} candidate, with 15 samples and three warmups per group.", '',
             'All ten allocation, work, read and output accounting sets match exactly across 68 control groups. Candidate-only date/time cases have no supported predecessor baseline.', '',
             'Median threshold flags are retained below. The machine-readable report includes every group and all median/p95/p99 flags and uncertainty intervals.', '',
             '| Group | Metric | Baseline | Candidate | Change | 95% bootstrap interval |',
             '| --- | --- | ---: | ---: | ---: | --- |']
    for flag in median:
        interval = flag['bootstrap']['relative_ci95']
        lines.append(f"| {flag['group']} | {flag['metric']} | {flag['baseline']:.4f} | {flag['candidate']:.4f} | {100*flag['relative']:+.3f}% | {100*interval['low']:+.3f}% to {100*interval['high']:+.3f}% |")
    lines += ['', *notes, '', 'Units: elapsed time is nanoseconds per repeat; RSS is KiB. Full environment, raw child JSON, time receipts and source/harness/lock identities are retained beside this report.', '']
    (RESULTS / 'performance-report.md').write_text('\n'.join(lines))
    write(RESULTS / 'retained-files.json', {
        str(path.relative_to(RESULTS)): digest(path)
        for path in sorted(RESULTS.rglob('*')) if path.is_file() and path.name != 'retained-files.json'
    })
    print(json.dumps({'status': 'verified', 'rows': audit['rows'], 'matched_groups': len(baseline), 'median_flags': len(median)}))


if __name__ == '__main__':
    main()
