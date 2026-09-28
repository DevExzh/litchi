"""Independent raw timing, semantic, and exact-owner frame cross-check."""
import collections
import gzip
import hashlib
import json
import math
from pathlib import Path
import random
import re
import statistics
import sys

P = Path(__file__).resolve().parent
OLD = P.parent / 'change-0813'
OWNER = 'namespace_uri_probe::capture_region_0793'

def read(path):
    return json.loads(path.read_text())

def bootstrap(values):
    rng = random.Random(814814)
    samples = sorted(statistics.median(values[rng.randrange(6)] for _ in range(6)) for _ in range(10000))
    return [samples[250], samples[9749]]

def analyze():
    seal = read(OLD / 'seal.json')['files']
    root = P.parents[3]
    native = read(P / 'native/receipts.json')
    perf = read(P / 'perf/receipts.json')
    assert len(native) == 54 and len(perf) == 2
    values = {}
    samples = 0
    for lane, receipts, expected_n in [('native', native, 30), ('perf', perf, 100)]:
        for receipt in receipts:
            path = Path(receipt['report']['path'])
            raw = path.read_bytes()
            assert hashlib.sha256(raw).hexdigest() == receipt['report']['sha256']
            report = json.loads(raw)
            assert report['schema'] == 'litchi.pptx.public-workflow-probe-0806.v1'
            assert report['mode'] == 'capture'
            shape = report['shape']
            oracle_path = OLD / f'native/0-{shape}-capture-after.json'
            assert hashlib.sha256(oracle_path.read_bytes()).hexdigest() == seal[str(oracle_path.relative_to(root))]
            oracle = read(oracle_path)
            for name in ('source', 'fixture', 'slides', 'shapes_per_slide', 'marker', 'timing_scope'):
                assert report[name] == oracle[name], (path, name)
            assert len(report['samples']) == expected_n
            assert report['warmup'] == (3 if lane == 'native' else 0)
            elapsed = []
            for index, item in enumerate(report['samples']):
                assert item['index'] == index
                for name in ('source_sha256', 'output', 'verification'):
                    assert item[name] == oracle['samples'][0][name], (path, index, name)
                assert item['elapsed_ns'] == item['metrics']['elapsed_ns'] > 0
                elapsed.append(item['elapsed_ns'])
            samples += len(elapsed)
            if lane == 'native':
                ordered = sorted(elapsed)
                key = (shape, receipt['variant'], receipt['block'])
                assert key not in values
                values[key] = {f'p{p}': ordered[math.ceil(len(ordered) * p / 100) - 1] for p in (50, 95, 99)}
    assert len(values) == 54 and samples == 1820
    comparisons = []
    for shape in ('tiny', 'medium', 'large'):
        for numerator, denominator in [('profile', 'control'), ('fp', 'profile')]:
            rows = []
            for block in range(6):
                before = values[(shape, denominator, block)]['p50']
                after = values[(shape, numerator, block)]['p50']
                rows.append({'block': block, 'before': before, 'after': after, 'ratio': after / before})
            ratios = [row['ratio'] for row in rows]
            comparisons.append({'shape': shape, 'comparison': f'{numerator}/{denominator}',
                                'blocks': rows, 'ratio': statistics.median(ratios), 'ci95': bootstrap(ratios)})
    frames = []
    paths = sorted((P / 'perf').glob('*.frames.gz'))
    assert len(paths) == 2, paths
    for path in paths:
        data = gzip.decompress(path.read_bytes()).decode()
        blocks = data.strip().split('\n\n')
        leaves, nested, occurrences = collections.Counter(), collections.Counter(), collections.Counter()
        qualified = unknown = period = 0
        for block in blocks:
            lines = block.splitlines()
            header = re.fullmatch(r'\S+\s+\d+\s+(\d+\.\d+):\s+(\d+) cycles:u:\s*', lines[0])
            assert header, lines[0]
            names = []
            for line in lines[1:]:
                match = re.fullmatch(r'\s*[0-9a-f]+ (.+?)(?:\+0x[0-9a-f]+)? \(.+\)', line)
                assert match, line
                names.append(match[1])
            if OWNER not in names:
                continue
            assert names.count(OWNER) == 1
            qualified += 1
            period += int(header[2])
            inside = names[:names.index(OWNER)]
            leaves[inside[0] if inside else OWNER] += 1
            nested.update(set(inside))
            occurrences.update(names[:names.index(OWNER) + 1])
            unknown += any(any(token in name.lower() for token in ('[unknown]', '??', '<unknown>')) for name in inside)
        assert sum(leaves.values()) == qualified
        frames.append({'path': path.name, 'whole_process_samples': len(blocks),
                       'owner_samples': qualified, 'owner_period': period, 'unknown_interior': unknown,
                       'leaf_counts': dict(sorted(leaves.items())), 'nested_counts': dict(sorted(nested.items())),
                       'frame_occurrences': dict(sorted(occurrences.items()))})
    return {'schema': 'litchi.performance.0814.root-audit.v1', 'reports': 56,
            'samples': samples, 'paired_p50': comparisons, 'frames': frames,
            'scope': 'Independent raw equality and sample counts; no native phase fraction or production speedup claim.'}

result = analyze()
encoded = json.dumps(result, indent=2, sort_keys=True) + '\n'
output = P / 'root-audit.json'
if '--check' in sys.argv:
    assert output.read_text() == encoded
else:
    assert not output.exists()
    output.write_text(encoded)
print('0814 independent raw audit PASS:56 reports/1820 samples')
