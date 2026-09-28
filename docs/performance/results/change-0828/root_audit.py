"""Independent raw native-report census and quantiles; no workload execution."""
import hashlib
import json
import math
from pathlib import Path
import random
import statistics as s
import sys

P = Path(__file__).resolve().parent


def read(path):
    return json.loads(path.read_text())


def checked(descriptor):
    path = Path(descriptor['path'])
    data = path.read_bytes()
    assert len(data) == descriptor['bytes']
    assert hashlib.sha256(data).hexdigest() == descriptor['sha256']
    return path


def main():
    assert sys.argv[1:] in (['--write'], ['--check'])
    assert (P / 'perf/decode-complete.json').is_file(), 'offline audit follows terminal decoding'
    plan = read(P / 'plan.json')
    groups = {arm: [] for arm in plan['arms']}
    reports = samples = 0
    for lane in ('qualification', 'native'):
        rows = read(P / lane / 'receipts.json')
        expected_order = [(block, arm) for block, order in enumerate(plan[lane]['orders'])
                          for arm in order]
        assert len(rows) == len(expected_order)
        for receipt, (block, arm) in zip(rows, expected_order, strict=True):
            assert receipt['label'] == f'{block:02d}-{arm}' and receipt['arm'] == arm
            assert receipt['exit_code'] == 0
            report = read(checked(receipt['report']))
            checked(receipt['log'])
            rss = int(checked(receipt['rss']).read_text().strip())
            assert report['all_verified'] and report['warmup_verified']
            assert report['mode'] == plan['arms'][arm]['mode']
            assert report['warmup'] == plan[lane]['warmup']
            records = report['samples']
            assert len(records) == plan[lane]['samples']
            raw = [row['elapsed_ns'] for row in records]
            assert report['elapsed_ns']['samples'] == raw
            assert report['elapsed_ns']['sample_order'] == list(range(len(raw)))
            assert all(x > 0 for x in raw)
            for index, row in enumerate(records):
                assert row['index'] == index and all(row['verification'].values())
                assert row['output'] == report['output']
                assert row['output']['bytes'] == 68284
                assert row['output']['sha256'] == '38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf'
            reports += 1
            samples += len(raw)
            if lane == 'native':
                ordered = sorted(raw)
                quantiles = {f'p{p}': ordered[math.ceil(p / 100 * len(raw)) - 1]
                             for p in (50, 95, 99)}
                groups[arm].append({'block': block, **quantiles,
                                    'mean': s.mean(raw), 'rss_kib': rss})
    assert (reports, samples) == (21, 549)
    medians = {arm: {metric: s.median(row[metric] for row in rows)
                     for metric in ('p50', 'p95', 'p99', 'mean', 'rss_kib')}
               for arm, rows in groups.items()}
    pairs = {}
    for numerator, denominator in (('wrapped', 'control'), ('fp', 'wrapped')):
        ratios = [n['p50'] / d['p50'] for n, d in
                  zip(groups[numerator], groups[denominator], strict=True)]
        rng = random.Random(828828)
        boot = sorted(s.median(rng.choice(ratios) for _ in ratios) for _ in range(10000))
        pairs[f'{numerator}/{denominator}'] = {
            'block_ratios': ratios, 'median': s.median(ratios),
            'ci95': [boot[250], boot[9749]],
        }
    result = {'schema': 'litchi.performance.0828.root-audit.v1',
              'unprofiled_reports': reports, 'unprofiled_samples': samples,
              'native_blocks': groups, 'native_medians': medians, 'paired_p50': pairs}
    encoded = json.dumps(result, indent=2, sort_keys=True) + '\n'
    out = P / 'root-audit.json'
    if sys.argv[1] == '--write':
        assert not out.exists()
        out.write_text(encoded)
    else:
        assert out.read_text() == encoded
    print('0828 independent native audit PASS: 21 reports, 549 samples')


if __name__ == '__main__':
    main()
