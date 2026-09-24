#!/usr/bin/env python3
"""Recheck every initial latency/RSS flag in four interleaved A/B pairs."""
import csv
import json
from pathlib import Path
import runpy
import statistics

PERF = Path(__file__).resolve().parent

def main():
    runner = runpy.run_path(str(PERF / 'logical-harness/run.py'))
    flags = json.loads((PERF / 'initial-flags.json').read_text())
    output = PERF / 'abab'
    output.mkdir(exist_ok=True)
    combined = []
    for round_number in range(1, 5):
        folder = output / f'r{round_number}'
        folder.mkdir(exist_ok=True)
        (folder / 'commands.txt').write_text('')
        rows, sequence = [], []
        for flag in flags:
            for revision in ['baseline', 'candidate']:
                row, item = runner['run_one'](
                    PERF / revision / 'ods-formula-logical-evaluation-profile', folder,
                    revision, 'comparable', flag['phase'], flag['case'], int(flag['repeat']), 3, 15)
                assert row['status'] == '0', row
                rows.append(row); sequence.append(item)
                combined.append(dict(round=round_number, **row))
        with (folder / 'raw.csv').open('w', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=runner['FIELDS'], lineterminator='\n')
            writer.writeheader(); writer.writerows(rows)
        (folder / 'sequence.json').write_text(json.dumps(sequence, indent=2) + '\n')
        print('round', round_number, len(rows), 'rows', flush=True)
    with (output / 'raw.csv').open('w', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=['round', *runner['FIELDS']], lineterminator='\n')
        writer.writeheader(); writer.writerows(combined)
    summary = []
    for flag in flags:
        selected = [r for r in combined if r['case'] == flag['case'] and r['phase'] == flag['phase']]
        item = dict(phase=flag['phase'], case=flag['case'], repeat=int(flag['repeat']))
        for metric in ['p50_ns', 'p95_ns', 'p99_ns', 'max_rss_kib', 'alloc_calls_p50', 'requested_bytes_p50', 'peak_live_delta_p50']:
            medians = {revision: statistics.median(float(r[metric]) for r in selected if r['revision'] == revision)
                       for revision in ['baseline', 'candidate']}
            item[metric] = dict(medians, delta_pct=(medians['candidate'] / medians['baseline'] - 1) * 100 if medians['baseline'] else 0)
        summary.append(item)
    (output / 'summary.json').write_text(json.dumps(summary, indent=2) + '\n')

if __name__ == '__main__':
    main()
