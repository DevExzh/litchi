"""Whole-child counter supplement, deliberately separate from method profiles."""
import csv
import json
from run import HERE

def analyze():
    rows = []
    expected = json.loads((HERE / 'plan.json').read_text())['hardware']['events']
    for path in sorted((HERE / 'hardware').glob('*.perf.csv')):
        counters = {}
        with path.open() as stream:
            for row in csv.reader(stream):
                if not row or row[0].startswith('#'):
                    continue
                assert len(row) >= 5 and float(row[4]) >= 99.9
                assert row[2] not in counters
                counters[row[2]] = dict(value=int(row[0]), runtime_ns=int(row[3]), coverage_percent=float(row[4]))
        assert set(counters) == set(expected)
        rows.append(dict(case=path.name.removesuffix('.perf.csv'), events=counters, IPC=counters['instructions']['value']/counters['cycles']['value']))
    assert len(rows) == 6
    return dict(scope='Whole child baseline only; includes fixture/preflight and output oracles; not publication attribution or candidate comparison', cases=rows)

if __name__ == '__main__':
    (HERE / 'hardware-summary.json').write_text(json.dumps(analyze(), indent=2)+'\n')
