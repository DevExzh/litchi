"""Independent raw native quantile and paired-bootstrap audit; no workloads."""
import json
import math
from pathlib import Path
import random
import statistics as s
import sys

P = Path(__file__).resolve().parent

def read(path):
    return json.loads(path.read_text())

def interval(values):
    rng = random.Random(821821)
    ordered = sorted(s.median([values[rng.randrange(6)] for _ in range(6)])
                     for _ in range(10000))
    return {'estimate': s.median(values), 'lower': ordered[250], 'upper': ordered[9749]}

def audit():
    plan = read(P / 'plan.json')
    analysis = read(P / 'analysis.json')
    assert len(analysis['native']) == len(plan['cases']) == 24
    rows = {}
    for case in plan['cases']:
        blocks = []
        for block in range(6):
            report = read(P / 'native' / f"{block:02}-{case['id']}.json")
            assert len(report['results']) == 1
            result = report['results'][0]
            assert result['case'] == case['case']
            samples = sorted(result['elapsed_ns']['samples'])
            assert len(samples) == 30 and all(x > 0 for x in samples)
            blocks.append({name: samples[math.ceil(30 * q) - 1]
                           for name, q in [('p50', .5), ('p95', .95), ('p99', .99)]}
                          | {'mean': s.mean(samples)})
        rows[case['id']] = blocks
    for row in analysis['native']:
        blocks = rows[row['id']]
        for metric in ('p50', 'p95', 'p99', 'mean'):
            values = [b[metric] for b in blocks]
            assert math.isclose(row[metric + '_ns'], s.median(values), rel_tol=1e-12)
            assert all(math.isclose(a, b, rel_tol=1e-12) for a, b in zip(row[metric + '_block_values_ns'], values))
            assert math.isclose(row['spread_ratios'][metric], max(values) / min(values), rel_tol=1e-12)
        p50 = [b['p50'] for b in blocks]
        absolute = interval(p50)
        assert all(row['p50_bootstrap_ci95_ns'][k] == v for k, v in absolute.items())
        default = rows[row['case'] + '__default']
        ratios = [b['p50'] / d['p50'] for b, d in zip(blocks, default)]
        paired = row['paired_ratio_to_default']
        assert paired['block_p50_ratios'] == ratios
        assert all(paired['bootstrap_ci95'][k] == v for k, v in interval(ratios).items())
        assert paired['ratio_median'] == s.median(ratios)
        assert row['tail_flag'] == (s.median([b['p99'] for b in blocks]) / s.median(p50) > 1.05)
    return {'schema': 'litchi.performance.0821.root-audit.v1', 'status': 'pass',
            'native_reports': 144, 'native_samples': 4320, 'rows': 24,
            'absolute_intervals': 24, 'paired_intervals': 24,
            'scope': 'independent raw quantiles, block spreads, tails, and paired bootstrap; custody validated separately'}

if __name__ == '__main__':
    assert sys.argv[1:] in (['--write'], ['--check'])
    value = audit()
    path = P / 'root-audit.json'
    if sys.argv[1] == '--write':
        assert not path.exists()
        path.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')
    else:
        assert read(path) == value
    print('0821 independent raw audit PASS: 144 reports, 4320 samples, 24 paired intervals')
