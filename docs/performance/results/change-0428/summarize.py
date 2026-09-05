#!/usr/bin/env python3
"""Derive phase distributions and same-build repeat review triggers from raw reports."""
import argparse
from collections import defaultdict
import json
from pathlib import Path
import statistics
import subprocess
import sys

ROOT = Path(__file__).resolve().parent


def numbers(value, prefix=''):
    for key, item in value.items():
        name = prefix + key
        if isinstance(item, dict):
            yield from numbers(item, name + '.')
        elif type(item) is int:
            yield name, item


def distribution(values):
    ordered = sorted(values)
    return {'count': len(values), 'min': ordered[0], 'median': statistics.median(ordered),
            'max': ordered[-1]}


def derive():
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    index = json.loads((ROOT / 'capture-index.json').read_text())
    processes = []
    comparisons = defaultdict(dict)
    retained = 0
    for name in index:
        receipt = json.loads((ROOT / name).read_text())
        path = ROOT / 'capture' / (receipt['name'] + '.json')
        subprocess.check_output([sys.executable, '-B', str(ROOT / 'verify-report.py'), str(path)])
        report = json.loads(path.read_text())
        fields = defaultdict(list)
        availability = defaultdict(set)
        for sample in report['samples_raw']:
            retained += 1
            labels = [point['label'] for point in sample['phases']]
            assert len(set(labels)) == len(labels), 'ambiguous phase labels'
            for point in sample['phases']:
                label = point['label']
                for key, value in numbers(point):
                    fields[(label, key)].append(value)
                for owner, value in point.items():
                    if isinstance(value, dict) and 'availability' in value:
                        availability[(label, owner)].add(value['availability'])
        rows = {}
        for (label, key), values in sorted(fields.items()):
            stats = distribution(values)
            rows.setdefault(label, {})[key] = stats
            comparisons[(receipt['scenario'], receipt['corpus'], label, key)][receipt['repeat']] = stats
        processes.append({'name': receipt['name'], 'scenario': receipt['scenario'],
                          'corpus': receipt['corpus'], 'repeat': receipt['repeat'],
                          'samples': len(report['samples_raw']), 'phases': rows,
                          'availability': {label + '.' + owner: sorted(values)
                                           for (label, owner), values in sorted(availability.items())}})
    review = []
    for (scenario, corpus, label, key), repeats in sorted(comparisons.items()):
        assert set(repeats) == {'R1', 'R2'}
        before, after = repeats['R1']['median'], repeats['R2']['median']
        change = 100 * (after - before) / before if before else None
        triggered = abs(change) > 5 if change is not None else after != 0
        if triggered:
            review.append({'scenario': scenario, 'corpus': corpus, 'phase': label, 'field': key,
                           'r1': repeats['R1'], 'r2': repeats['R2'], 'median_change_percent': change,
                           'scope': 'process RSS point' if key.startswith('rss.') else 'managed/source observation'})
    assert len(processes) == protocol['expected_processes']
    assert retained == protocol['expected_retained_samples']
    return {'classification': 'descriptive same-build phase distributions; no optimization comparison',
            'processes': processes, 'retained_samples': retained,
            'repeat_review_threshold_percent': 5, 'repeat_review_triggers': review,
            'performance_claim': None,
            'scope': 'Cache/budget gauges, cumulative charges, source reads and process RSS remain separate; null observations are not zero.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = derive()
    target = ROOT / 'summary.json'
    if args.check:
        assert json.loads(target.read_text()) == result, 'summary differs from raw reports'
    else:
        assert not target.exists()
        target.write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'samples': result['retained_samples'],
                      'repeat_review_triggers': len(result['repeat_review_triggers'])}))


if __name__ == '__main__':
    main()
