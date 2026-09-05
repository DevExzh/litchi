#!/usr/bin/env python3
"""Derive descriptive checkpoint tables from validated raw reports."""
import argparse
import json
from pathlib import Path
import statistics
import subprocess
import sys

ROOT = Path(__file__).resolve().parent


def stats(values):
    return {'count': len(values), 'min': min(values), 'max': max(values),
            'mean': statistics.mean(values), 'median': statistics.median(values),
            'sum': sum(values)}


def derive(root):
    index = json.loads((root / 'capture-index.json').read_text())
    runs = []
    for name in index:
        receipt = json.loads((root / name).read_text())
        report_path = root / 'capture' / (receipt['name'] + '.json')
        subprocess.check_output([sys.executable, '-B', str(root / 'verify-report.py'), str(report_path)])
        report = json.loads(report_path.read_text())
        rows = report['samples_raw']
        phases = [phase['label'] for phase in report['phases']]
        stages = []
        previous = phases[0]
        for phase in phases:
            stages.append({
                'phase': phase,
                'absolute_live_bytes': stats([row[phase]['live_bytes'] for row in rows]),
                'live_change_from_entry_bytes': stats([
                    row[phase]['live_bytes'] - row[phases[0]]['live_bytes'] for row in rows]),
                'live_change_from_previous_bytes': stats([
                    row[phase]['live_bytes'] - row[previous]['live_bytes'] for row in rows]),
                'allocation_requests_from_entry_bytes': stats([
                    row[phase]['allocated_bytes'] - row[phases[0]]['allocated_bytes'] for row in rows]),
                'deallocation_requests_from_entry_bytes': stats([
                    row[phase]['deallocated_bytes'] - row[phases[0]]['deallocated_bytes'] for row in rows]),
            })
            previous = phase
        runs.append({
            'name': receipt['name'], 'api': report['api'], 'corpus': report['corpus'],
            'repeat': receipt['repeat'], 'stages': stages,
            'whole_probe_region_peak_live_bytes': stats([
                row['retention_probe']['retention_probe_region_peak_live_bytes'] for row in rows]),
            'whole_probe_peak_above_entry_bytes': stats([
                row['retention_probe']['retention_probe_region_peak_live_bytes']
                - row[phases[0]]['live_bytes'] for row in rows]),
            'source_archive_sha256': report['source_archive_sha256'],
            'destination_archive_sha256': report['destination_archive_sha256'],
            'expected_output_sha256': report['expected_output_sha256'],
        })
    drift = []
    for corpus in ['plain', 'media-rich']:
        for api in ['owned', 'source-backed']:
            pair = sorted([run for run in runs if run['api'] == api and run['corpus'] == corpus],
                          key=lambda run: run['repeat'])
            assert len(pair) == 2
            for first, second in zip(pair[0]['stages'], pair[1]['stages'], strict=True):
                assert first['phase'] == second['phase']
                left = first['live_change_from_entry_bytes']['mean']
                right = second['live_change_from_entry_bytes']['mean']
                percent = 100 * (right - left) / abs(left) if left else (0 if right == 0 else None)
                drift.append({'api': api, 'corpus': corpus, 'phase': first['phase'],
                              'r1_mean_change_bytes': left, 'r2_mean_change_bytes': right,
                              'repeat_change_percent': percent,
                              'review_required': percent is None or abs(percent) > 5})
    return {'classification': 'descriptive callback-order allocator checkpoints',
            'processes': len(runs), 'runs': runs, 'repeat_drift': drift,
            'performance_claim': None}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = derive(ROOT)
    path = ROOT / 'summary.json'
    raw = json.dumps(result, indent=2, sort_keys=True) + '\n'
    if args.check:
        assert json.loads(path.read_text()) == result
        print(json.dumps({'status': 'pass', 'processes': len(result['runs']), 'summary_recomputed': True}))
        return
    assert not path.exists()
    path.write_text(raw)
    lines = ['# Descriptive live-byte changes from sample entry', '',
             'All table values are mean MiB above each sample’s entry checkpoint.',
             'They are global allocator observations, not object-owned memory or RSS.', '',
             '| Corpus | API | Repeat | Planned | Published | Plan dropped | Documents dropped | Sink dropped |',
             '| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |']
    for run in result['runs']:
        stages = {stage['phase']: stage for stage in run['stages']}
        values = [stages[name]['live_change_from_entry_bytes']['mean'] / (1024 * 1024)
                  for name in ['planned', 'published', 'drop_plan', 'drop_document_handles', 'drop_sink']]
        lines.append('| ' + ' | '.join([run['corpus'], run['api'], run['repeat'],
                                      *[f'{value:.6f}' for value in values]]) + ' |')
    (ROOT / 'result-table.md').write_text('\n'.join(lines) + '\n')
    print(json.dumps({'status': 'pass', 'processes': len(result['runs']),
                      'drift_review_rows': sum(row['review_required'] for row in result['repeat_drift'])}))


if __name__ == '__main__':
    main()
