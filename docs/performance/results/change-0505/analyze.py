#!/usr/bin/env python3
"""Validate and summarize retained 0505 timing pairs; no benchmark execution."""
import argparse
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location('odg_summary', ROOT.parent / 'change-0502' / 'summarize.py')
SUMMARY = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(SUMMARY)


def rss(path):
    prefix = 'Maximum resident set size (kbytes):'
    for line in path.read_text().splitlines():
        if line.strip().startswith(prefix):
            return int(line.split(':', 1)[1])
    raise ValueError(f'missing RSS: {path}')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=ROOT)
    root = parser.parse_args().root
    result = SUMMARY.summarize(root)
    result['schema'] = 'litchi.odg.attribute-selection-summary.v1'
    result['scope'] = 'same semantics; owned open plus traversal; A1/B1/B2/A2 serial blocks'
    flags = []
    for row in result['rows']:
        corpus, repeat = row['corpus'], row['repeat']
        paths = {phase: root / phase / f'{corpus}-{repeat}.json' for phase in ('before', 'after')}
        reports = {phase: json.loads(path.read_text()) for phase, path in paths.items()}
        for report in reports.values():
            values = report['elapsed_ns']
            expected = {
                'min_ns': min(values), 'max_ns': max(values),
                'mean_ns': sum(values) / len(values),
                **{f'p{percent}_ns': SUMMARY.probe_percentile(values, percent / 100)
                   for percent in (50, 95, 99)},
            }
            expected['throughput_input_bytes_per_second_p50'] = report['input_bytes'] * 1_000_000_000 / expected['p50_ns']
            for metric, value in expected.items():
                assert abs(report['statistics'][metric] / value - 1) < 1e-12, (corpus, repeat, metric)
        assert reports['before']['semantic_checksum'] == reports['after']['semantic_checksum']
        row['semantic_checksum'] = reports['before']['semantic_checksum']
        row['statistics'] = {}
        for metric in ('p50_ns', 'p95_ns', 'p99_ns', 'mean_ns', 'throughput_input_bytes_per_second_p50'):
            before, after = (reports[phase]['statistics'][metric] for phase in ('before', 'after'))
            ratio = after / before - 1
            row['statistics'][metric] = {'before': before, 'after': after, 'change_ratio': ratio}
            adverse = ratio < -0.05 if metric.startswith('throughput') else ratio > 0.05
            if adverse:
                flags.append({'corpus': corpus, 'repeat': repeat, 'metric': metric, 'change_ratio': ratio})
        before, after = (rss(paths[phase].with_suffix('.time.txt')) for phase in ('before', 'after'))
        row['whole_child_rss_kib'] = {'before': before, 'after': after, 'change_ratio': after / before - 1}
        if after / before > 1.05:
            flags.append({'corpus': corpus, 'repeat': repeat, 'metric': 'whole_child_rss_kib', 'change_ratio': after / before - 1})
    result['adverse_over_five_percent'] = flags
    (root / 'summary.json').write_text(json.dumps(result, indent=2, sort_keys=True) + '\n')
    for row in result['rows']:
        print(row['corpus'], row['repeat'], f"p50 {row['p50_change_ratio']:+.2%}", f"RSS {row['whole_child_rss_kib']['change_ratio']:+.2%}")
    print(f'{len(flags)} adverse >5% flags retained')


if __name__ == '__main__':
    main()
