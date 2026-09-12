"""Negative custody probes using in-memory tampering; raw evidence is untouched."""
import copy
import datetime
import json
from unittest.mock import patch
import verify as V


def must_reject(name, action):
    try:
        action()
    except V.VerificationError as error:
        return dict(name=name, status='pass', observed_refusal=str(error))
    raise AssertionError('tampered evidence was accepted: ' + name)


def run_tests():
    results = []
    plan = V.check_plan()
    folder = V.HERE / 'baseline'
    name = 'native-r1-xls'
    receipt = V.read_json(folder / (name + '.receipt.json'))
    expected = V.expected_artifacts(folder, name, 'native')
    V.validate_artifacts(folder, receipt, expected, 'positive control')
    for label, mutate in [
        ('omitted-host-artifact', lambda r: r['artifacts'].pop(name + '.host.json')),
        ('forged-raw-vector-hash', lambda r: r['artifacts'].__setitem__(name + '.json', '0' * 64)),
        ('unexpected-extra-artifact', lambda r: r['artifacts'].__setitem__('extra.json', '0' * 64)),
    ]:
        changed = copy.deepcopy(receipt)
        mutate(changed)
        results.append(must_reject(label, lambda: V.validate_artifacts(folder, changed, expected, label)))
    now = datetime.datetime(2026, 9, 12, tzinfo=datetime.timezone.utc)
    intervals = [(now, now + datetime.timedelta(seconds=2)),
                 (now + datetime.timedelta(seconds=1), now + datetime.timedelta(seconds=3))]
    results.append(must_reject('overlapping-capture-intervals', lambda: V.check_serial(intervals, 'probe')))
    quality, quality_intervals = V.validate_quality(plan)
    optional_intervals = V.validate_optional_quality_receipts(plan)
    assert len(quality_intervals) == quality["checks"] == 14
    assert not optional_intervals, "selected quality stage was counted twice"
    original_read = V.read_json
    changed = copy.deepcopy(plan)
    changed['review']['primary_cases'] = changed['review']['primary_cases'][:1]
    def altered_plan(path, *args, **kwargs):
        return changed if path == V.PLAN else original_read(path, *args, **kwargs)
    with patch.object(V, 'read_json', side_effect=altered_plan):
        results.append(must_reject('silently-narrowed-primary-gate', V.check_plan))
    return dict(status='pass', tests=results,
        raw_receipt_sha256=V.sha(folder / (name + '.receipt.json')),
        verifier_sha256=V.sha(V.HERE / 'verify.py'),
        scope='Five in-memory negative custody probes plus raw-artifact and non-duplicated quality-interval positive controls. No Rust test or benchmark execution.')


if __name__ == '__main__':
    result = run_tests()
    (V.HERE / 'verifier-tests.json').write_text(json.dumps(result, indent=2) + '\n')
    print('Five negative custody probes passed')
