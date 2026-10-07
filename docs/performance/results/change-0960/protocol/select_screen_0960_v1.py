#!/usr/bin/env python3
"""Apply the pre-capture root selection rule; never authorize adoption."""
import argparse
import hashlib
import json
import math
from pathlib import Path

CASES = {'medium-one-percent', 'dense-sparse-one-percent', 'medium-one-edit'}
PRIMARY = ('edit_commit_ns.p50', 'edit_commit_ns.p95', 'total_ns.p50', 'total_ns.p95')

def select(report, rule):
    assert report['status'] == 'screened'
    assert rule['signal_threshold'] == 0.95 and rule['safety_review_threshold'] == 1.05
    stops, signals, observed = [], [], {}
    for lane, expected in [('native', 6), ('allocator', 3)]:
        assert set(report[lane]['cases']) == CASES
        observed[lane] = {}
        for case, value in report[lane]['cases'].items():
            metrics = value['metrics']
            required = (*PRIMARY, 'max_rss_kib') if lane == 'native' else ('allocated_bytes', 'allocation_calls', 'region_peak_live_bytes', 'core_combined_with_drain.allocated_bytes', 'max_rss_kib')
            observed[lane][case] = {}
            for key in required:
                metric = metrics.get(key, {})
                ratio = metric.get('median')
                label = f'{lane}/{case}/{key}'
                valid = (metric.get('pair_count') == expected and metric.get('undefined_count') == 0
                         and type(ratio) in (float, int) and math.isfinite(ratio) and ratio > 0)
                observed[lane][case][key] = metric
                if not valid:
                    stops.append(label + ': invalid or undefined')
                    continue
                if lane == 'native' and key in PRIMARY and ratio > 1.05:
                    stops.append(label + ': primary median exceeds 1.05')
                eligible = lane == 'native' or key == 'region_peak_live_bytes'
                if eligible and ratio <= 0.95:
                    signals.append({'metric': label, 'median': ratio})
    if not signals:
        stops.append('No qualifying metric among signal_any')
    return {'decision': 'reject_screen' if stops else 'require_full_confirmation_and_regression_review',
            'adoption_allowed': False, 'statistical_effect_claim': False,
            'stops': stops, 'signals': signals, 'required_metrics': observed}

def descriptor(path):
    data = path.read_bytes()
    return {'path': str(path.resolve()), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    for key in ('report', 'rule', 'output'):
        parser.add_argument('--' + key, required=True, type=Path)
    args = parser.parse_args()
    result = select(json.loads(args.report.read_text()), json.loads(args.rule.read_text()))
    result.update(schema='litchi.performance.0960.root-screen-selection.v1',
                  report=descriptor(args.report), rule=descriptor(args.rule), helper=descriptor(Path(__file__)))
    with args.output.open('x') as output:
        json.dump(result, output, indent=2)
        output.write('\n')
    print(result['decision'])
