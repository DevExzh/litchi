#!/usr/bin/env python3
"""Derive API-only duration distributions and retain every repeat review flag."""
import argparse
import hashlib
import json
from pathlib import Path
import random
import statistics

ROOT = Path(__file__).resolve().parent


def flatten(value, prefix=''):
    result = {}
    for key, item in value.items():
        name = prefix + key
        if isinstance(item, dict):
            result.update(flatten(item, name + '.'))
        elif type(item) is int:
            result[name] = item
        elif isinstance(item, list) and all(type(entry) is int for entry in item):
            result.update({name + '.' + str(index): entry for index, entry in enumerate(item)})
    return result


def percentile(values, fraction):
    ordered = sorted(values)
    position = (len(ordered) - 1) * fraction
    left = int(position)
    right = min(left + 1, len(ordered) - 1)
    return ordered[left] + (ordered[right] - ordered[left]) * (position - left)


def duration_summary(values, label):
    rng = random.Random(int.from_bytes(hashlib.sha256(label.encode()).digest()[:8], 'big'))
    boot = [statistics.median([values[rng.randrange(len(values))] for _ in values]) for _ in range(2000)]
    return {'samples': len(values), 'min': min(values), 'p50': percentile(values, .5),
            'p95': percentile(values, .95), 'p99': percentile(values, .99), 'max': max(values),
            'median_bootstrap_95_percentile_interval': [percentile(boot, .025), percentile(boot, .975)]}


def derive():
    captures = []
    groups = {}
    sample_count = 0
    phase_count = 0
    for path in json.loads((ROOT / 'capture-index.json').read_text()):
        receipt = json.loads((ROOT / path).read_text())
        report = json.loads((ROOT / 'capture' / (receipt['name'] + '.json')).read_text())
        rows = report['samples_raw']
        sample_count += len(rows)
        metrics = {}
        result_bytes = report['expected_output_bytes'] if receipt['command'] == 'provider-lifecycle' else report['payload_oracle']['payload_bytes']
        for row in rows:
            assert row['timings']['api_sum_ns'] > 0
            rate = 1_000_000_000 / row['timings']['api_sum_ns']
            metrics.setdefault('rates.completed_operations_per_api_second', []).append(rate)
            metrics.setdefault('rates.logical_result_bytes_per_api_second', []).append(rate * result_bytes)
            for key, value in row['timings'].items():
                assert type(value) is int and value >= 0
                metrics.setdefault('timings.' + key, []).append(value)
            for phase in row['phases']:
                phase_count += 1
                for key, value in flatten(phase).items():
                    metrics.setdefault('phases.' + phase['label'] + '.' + key, []).append(value)
        distributions = {}
        for key, values in sorted(metrics.items()):
            if key.startswith(('timings.', 'rates.')):
                distributions[key] = duration_summary(values, receipt['name'] + '/' + key)
            else:
                distributions[key] = {'samples': len(values), 'min': min(values), 'p50': statistics.median(values), 'max': max(values)}
        row = {'name': receipt['name'], 'command': receipt['command'], 'selector': receipt['selector'],
               'provider_label': receipt['provider_label'], 'repeat': receipt['repeat'], 'metrics': distributions}
        captures.append(row)
        role = (row['command'], row['selector'], row['provider_label'])
        assert row['repeat'] not in groups.setdefault(role, {})
        groups[role][row['repeat']] = row
    flags = []
    for role, pair in sorted(groups.items()):
        before, after = pair['R1']['metrics'], pair['R2']['metrics']
        assert set(before) == set(after)
        for metric in sorted(before):
            a, b = before[metric]['p50'], after[metric]['p50']
            if a == b:
                continue
            percent = None if a == 0 else 100 * (b - a) / a
            if percent is None or abs(percent) > 5:
                flags.append({'command': role[0], 'selector': role[1], 'provider_label': role[2],
                              'metric': metric, 'r1_p50': a, 'r2_p50': b, 'change_percent': percent,
                              'scope': 'API duration point' if metric.startswith('timings.') else
                                       'API-derived rate' if metric.startswith('rates.') else
                                       'process RSS point' if '.rss.' in metric else 'managed resource or logical read point'})
    return {'status': 'pass', 'processes': len(captures), 'retained_samples': sample_count,
            'phase_points': phase_count,
            'method': 'p50/p95/p99 use linear interpolation of each 30-sample process; 2000 resamples with replacement and SHA256-derived fixed seeds yield descriptive percentile bootstrap median intervals. Small sample tails and dependent machine state limit population inference.',
            'timing_scope': 'Only named API calls; their sum excludes setup, source construction, reservations, checks, observers and drops. Rates divide one completed operation and its output archive bytes (cross-copy) or selected payload bytes (native image) by this API-only sum. No full lifecycle or causal optimization claim.',
            'captures': captures, 'repeat_review_triggers': flags, 'performance_claim': None}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = derive()
    target = ROOT / 'summary.json'
    if args.check:
        assert json.loads(target.read_text()) == result
    else:
        assert not target.exists()
        target.write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps({key: result[key] for key in ['status', 'processes', 'retained_samples', 'phase_points']} | {'repeat_flags': len(result['repeat_review_triggers'])}))


if __name__ == '__main__':
    main()
