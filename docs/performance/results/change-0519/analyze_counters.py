"""Compare owner/source gauges without treating them as allocator metrics."""
import csv
import json
from run import HERE

EQUAL_FIELDS = ['source_read_calls', 'source_requested_bytes', 'source_returned_bytes', 'output_bytes', 'output_sha256', 'budget_live_memory', 'budget_live_objects', 'budget_after_memory', 'budget_after_objects', 'budget_after_input', 'budget_after_output']

def analyze():
    cases = []
    total = 0
    for lane in ['r1', 'r2']:
        paths = sorted((HERE / lane).glob('*.csv'))
        assert len(paths) == 24
        for path in paths:
            with path.open() as stream:
                left = list(csv.DictReader(stream))
            with (HERE / ('after-' + lane) / path.name).open() as stream:
                right = list(csv.DictReader(stream))
            assert len(left) == len(right) == 66
            for a, b in zip(left, right):
                for key in ['repeat', 'ordinal', 'warmup', *EQUAL_FIELDS]:
                    assert a[key] == b[key], (lane, path.name, key)
            deltas = {int(a['budget_after_work']) - int(b['budget_after_work']) for a, b in zip(left, right)}
            assert deltas == {0}, (lane, path.name, deltas)
            cases.append(dict(lane=lane, case=path.stem, removed_work_bytes=deltas.pop()))
            total += len(left)
    return dict(equal_fields=EQUAL_FIELDS, rows_checked=total, cases=cases, scope='Owner gauges and source counters; not allocator metrics')

if __name__ == '__main__':
    (HERE / 'counter-comparison.json').write_text(json.dumps(analyze(), indent=2)+'\n')
