"""Independent baseline-corpus and source-counter acceptance before timing."""
import json
import math
import sys
import custody as c

p = c.P
old = p.parent / 'change-0786'
accepting = sys.argv[1:] == ['--write']
assert accepting or sys.argv[1:] == ['--check']
if accepting:
    assert not (p / 'native').exists()
rows = c.read(p / 'qualification/receipts.json')
assert len(rows) == 72
seal = c.read(old / 'seal.json')['files']
references = {}
for shape in ('large', 'mixed'):
    path = old / f'qualification/0-cfb-{shape}-fresh-65536-1.json'
    key = str(path.relative_to(old))
    assert seal[key] == c.sha(path)
    references[shape] = {'artifact': c.artifact(path), 'report': c.read(path)}
checks = []
for row in rows:
    report = c.read(row['report']['path'])
    assert c.artifact(row['report']['path']) == row['report']
    ref = references[row['shape']]['report']
    assert report['corpus'] == ref['corpus']
    assert len(report['samples']) == 1
    sample = report['samples'][0]
    assert sample['verification'] == ref['samples'][0]['verification']
    metrics = sample['source_metrics']
    resources = sample['resources']
    assert resources['worker_and_io_released'] and resources['cpu_tasks_within_limit']
    for stage in ('before_operation', 'after_operation', 'after_drop'):
        assert 0 <= resources[stage]['workers'] <= row['workers']
        assert 0 <= resources[stage]['io_concurrency'] <= row['workers']
        assert resources[stage]['cpu_tasks'] <= 1_000_000
    assert resources['after_drop']['workers'] == resources['after_drop']['io_concurrency'] == 0
    assert metrics['active_reads_after_operation'] == 0
    assert 0 <= metrics['max_simultaneous_reads'] <= row['workers']
    assert sum(metrics['request_size_histogram']) == metrics['logical_calls']
    if row['route'] == 'cfb':
        sizes = [member['bytes'] for member in report['corpus']['members']]
        cap = row['source_max_read_bytes']
        calls = sum(math.ceil(size / cap) if cap else 1 for size in sizes)
        requested = sum(sum(range(size, 0, -cap)) if cap else size for size in sizes)
        expected = {'logical_calls': calls, 'returned_bytes': sum(sizes),
                    'requested_bytes': requested, 'short_reads': calls - 32}
    elif row['state'] == 'primed':
        expected = dict.fromkeys(('logical_calls', 'returned_bytes', 'requested_bytes', 'short_reads'), 0)
        assert metrics['max_simultaneous_reads'] == 0
        assert sample['primed_cache_hit_control']
    else:
        count = 146041 if row['shape'] == 'large' else 143782
        expected = {'logical_calls': 64, 'returned_bytes': count,
                    'requested_bytes': count, 'short_reads': 0}
    assert all(metrics[k] == v for k, v in expected.items()), (row, metrics, expected)
    checks.append({'report': row['report'], 'expected_source': expected,
                   'max_simultaneous_reads': metrics['max_simultaneous_reads']})
result = {'schema': 'litchi.performance.0816.qualification-audit.v1',
          'reports': 72, 'samples': 72, 'accepted_before_native': True,
          'timings_imported': False,
          'references': {k: v['artifact'] for k, v in references.items()}, 'checks': checks}
path = p / 'qualification-audit.json'
encoded = json.dumps(result, indent=2, sort_keys=True) + '\n'
if accepting:
    assert not path.exists()
    path.write_text(encoded)
else:
    assert path.read_text() == encoded
print('0816 independent qualification corpus and source-counter audit PASS')
