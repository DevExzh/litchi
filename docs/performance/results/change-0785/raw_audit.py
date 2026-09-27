"""Independent raw-report decision cross-check; no benchmark execution."""
import json
from pathlib import Path
import random
import statistics
import sys

P = Path(__file__).resolve().parent

def read(path):
    return json.loads(path.read_text())

def emit(name, value):
    text = json.dumps(value, indent=2) + '\n'
    path = P / name
    if '--check' in sys.argv:
        assert path.read_text() == text, name
    else:
        path.write_text(text)

plan = read(P / 'plan.json')
rows = []
differences = []
paired_samples = 0
for case in plan['cases']:
    shape, mode = case['shape'], case['mode']
    ratios = []
    for block in range(6):
        values = {}
        for leg in ('before', 'after'):
            report = read(P / f'native/{block}-{shape}-{mode}-{leg}.json')
            assert len(report['samples']) == 30
            assert all(x['verification']['semantic_check'] for x in report['samples'])
            values[leg] = sorted(x['elapsed_ns'] for x in report['samples'])[14]
        ratios.append(values['after'] / values['before'])
    rng = random.Random(785078)
    boot = sorted(statistics.median(rng.choices(ratios, k=6)) for _ in range(10000))
    ratio = statistics.median(ratios)
    rows.append({**case, 'ratio': ratio, 'percent': (ratio - 1) * 100,
                 'ci': [boot[249], boot[9749]], 'pairs': ratios})
    for block in range(2):
        before = read(P / f'allocation/{block}-{shape}-{mode}-before.json')
        after = read(P / f'allocation/{block}-{shape}-{mode}-after.json')
        assert len(before['samples']) == len(after['samples']) == 3
        for index, (x, y) in enumerate(zip(before['samples'], after['samples'])):
            values = []
            for v in (x['allocation'], y['allocation']):
                assert v['status'] == 'measured'
                values.append({**v, 'net_live': v['live_bytes_after'] - v['live_bytes_before'],
                               'peak_above_entry': v['region_peak_live_bytes'] - v['live_bytes_before']})
            for key in ('net_live', 'peak_above_entry', 'allocation_calls', 'allocated_bytes'):
                if values[0][key] != values[1][key]:
                    differences.append([shape, mode, block, index, key, values[0][key], values[1][key]])
            paired_samples += 1
emit('root-raw-latency-audit.json', {
    'scope': 'Independent direct raw-report nearest-rank processp50 audit; seededbootstrapmedianpairedratios.',
    'rows': rows})
emit('root-raw-memory-audit.json', {'paired_samples': paired_samples,
    'fields': ['net_live', 'peak_above_entry', 'allocation_calls', 'allocated_bytes'],
    'differences': differences})
print('Independent latency and memory audit matches.' if '--check' in sys.argv else 'Wrote independent raw audit.')
