"""Replay archived gate receipts and packet integrity without Cargo/native work."""
import hashlib
import json
from pathlib import Path
import re
import tempfile
import verify_trace

P = Path(__file__).resolve().parent


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def validate():
    results = {}
    for name, count in [('quality', 8), ('broader', 4)]:
        receipt = json.loads((P / (name + '.json')).read_text())
        assert len(receipt['rows']) == count, name
        assert sha(P / receipt['source_file']) == receipt.get('source_sha256'), name
        rows = []
        for row in receipt['rows']:
            assert row['exit'] == 0, row
            log = P / row['log']
            assert sha(log) == row['sha256'], log
            matches = re.findall(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored', log.read_text())
            assert all(x[0] == 'ok' and x[2] == '0' for x in matches), log
            rows.append({'log': row['log'], 'suites': len(matches),
                         **{key: sum(int(x[i]) for x in matches)
                            for key, i in [('passed', 1), ('failed', 2), ('ignored', 3)]}})
        results[name] = rows
    assert results['quality'][2]['passed'] == 1738
    assert results['quality'][3]['passed'] == 8
    for row, expected in zip(results['broader'][:3], [10466, 823, 382]):
        assert row['passed'] == expected, row
    assert results['broader'][3]['passed'] == 555
    extra = P / 'extra-gates.json'
    if extra.exists():
        for row in json.loads(extra.read_text())['rows']:
            assert row['exit'] == 0 and sha(P / row['log']) == row['sha256'], row
    trace = P / 'trace-0'
    run = json.loads((trace / 'run.json').read_text())
    for leg in ['before', 'after']:
        for key in ['trace', 'stdout', 'stderr']:
            assert sha(trace / run['legs'][leg][key]) == run['legs'][leg][key + '_sha256']
    with tempfile.TemporaryDirectory(prefix='litchi-0773-offline-') as temporary:
        report = verify_trace.verify(trace / 'before.strace', trace / 'after.strace',
                                     trace / 'run.json', Path(temporary) / 'windows.json')
        assert report['ok'], report.get('errors')
    results['trace'] = {key: report[key] for key in ['ok', 'window_counts',
                        'level_expectations_hold', 'all_output_bytes_identical',
                        'all_levels_are_exact_subsequences']}
    seal = P / 'seal.json'
    if seal.exists():
        manifest = json.loads(seal.read_text())
        actual = {str(f.relative_to(P)): sha(f) for f in sorted(P.rglob('*'))
                  if f.is_file() and f != seal and '__pycache__' not in f.parts}
        assert actual == manifest['files'], 'packet differs from seal'
    return results


if __name__ == '__main__':
    print(json.dumps(validate(), indent=2))
