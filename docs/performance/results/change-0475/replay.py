#!/usr/bin/env python3
"""Recompute retained CPU and heap attribution from portable exports."""
import argparse
import json
from pathlib import Path
import cpu_analyze
import heap_analyze
import owners

ROOT = Path(__file__).resolve().parent


def equal(actual, expected, label):
    if actual != expected:
        raise ValueError(f'{label} attribution does not replay exactly')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cpu-only', action='store_true')
    args = parser.parse_args()
    for lane in ('cpu-P1', 'cpu-P2'):
        expected = json.loads((ROOT / f'{lane}-attribution.json').read_text())
        reports = [ROOT / item['file']['path'] for item in expected['inputs']['normal_reports']]
        actual = cpu_analyze.analyze_script(ROOT / lane / 'perf-script.stdout.gz', reports,
            profile_label=lane, bundle_root=ROOT)
        equal(actual, expected, lane)
        print(lane, 'exact replay passed', flush=True)
    if not args.cpu_only:
        equal(owners.analyze(ROOT), json.loads((ROOT / 'owners.json').read_text()), 'compressor owners')
        expected = json.loads((ROOT / 'heap-attribution.json').read_text())
        for trace in expected['traces']:
            lane = trace['lane']
            # Relative input names are part of the portable recorded identity.
            actual = heap_analyze.analyze_trace(Path(lane) / 'decoded.stdout.gz', lane,
                print_path=Path(lane) / 'print.stdout.gz').as_dict()
            equal(actual, trace, lane)
            print(lane, 'exact replay passed', flush=True)


if __name__ == '__main__':
    main()
