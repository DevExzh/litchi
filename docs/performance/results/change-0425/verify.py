#!/usr/bin/env python3
"""Replay retained verification custody and composite test coverage (no Cargo)."""
import gzip
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent


def raw_log(record):
    path = ROOT / record['log']['path']
    raw = path.read_bytes() if path.exists() else gzip.decompress(
        path.with_suffix(path.suffix + '.gz').read_bytes())
    assert len(raw) == record['log']['bytes'], path
    assert hashlib.sha256(raw).hexdigest() == record['log']['sha256'], path
    return raw


def main():
    expected = json.loads((ROOT / 'expected-checks.json').read_text())
    receipts = {}
    for name, status in expected.items():
        record = json.loads((ROOT / 'checks' / (name + '.json')).read_text())
        assert record['status'] == status, name
        assert record['source_unchanged'], name
        assert record['source_before'] == record['source_after'], name
        assert (record['exit_code'] == 0) == (status == 'pass'), name
        raw_log(record)
        receipts[name] = record
    actual = {p.stem for p in (ROOT / 'checks').glob('*.json')
              if 'argv' in json.loads(p.read_text())}
    assert actual == set(expected), 'unclassified or missing command receipt'

    full = receipts['non-iwork-full-tests-final']
    summaries = re.findall(
        rb'test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',
        raw_log(full))
    assert len(summaries) == 799
    assert sum(int(row[1]) for row in summaries) == 15775
    assert sum(int(row[2]) for row in summaries) == 3
    assert sum(int(row[3]) for row in summaries) == 98
    assert sum(row[0] == b'FAILED' for row in summaries) == 3

    corrected_files = {
        'crates/litchi-odp/tests/source_backed.rs',
        'crates/litchi-odt/tests/generic_content_publication.rs',
        'crates/litchi-xlsb/tests/workbook_structure_edit.rs',
    }
    corrected_count = 0
    for package, count in [('odp', 12), ('odt', 6), ('xlsb', 9)]:
        baseline = receipts['baseline-' + package + '-locked']
        assert b'test result: FAILED. 0 passed; 1 failed;' in raw_log(baseline)
        assert baseline['revision'] == '340cc91ae2bdec338dfe7682b4b5d8c219a2d288'
        candidate = receipts['corrected-' + package + '-tests']
        assert candidate['passed_tests'] == count
        before, after = full['source_after'], candidate['source_after']
        assert set(before) <= set(after)
        assert set(after) - set(before) == {'Cargo.lock'}
        changed = {name for name in before if before[name] != after[name]}
        assert changed == corrected_files, changed
        assert candidate['source_after']['Cargo.lock'] == baseline['source_after']['Cargo.lock']
        corrected_count += count
    # Replace the three failed target summaries, which already contained 24 passes.
    total = 15775 - 24 + corrected_count
    assert total == 15778
    facade_failures = []
    for name, expected_passes in [('facade-test', 455), ('baseline-facade-tests', 360)]:
        raw = raw_log(receipts[name])
        rows = re.findall(rb'test result: (ok|FAILED)\. (\d+) passed; (\d+) failed;', raw)
        assert sum(int(row[1]) for row in rows) == expected_passes
        assert sum(int(row[2]) for row in rows) == 6
        facade_failures.append(re.findall(rb'^test (.+) \.\.\. FAILED$', raw, re.M))
    assert len(facade_failures[0]) == 6 and facade_failures[0] == facade_failures[1]
    assert receipts['perf-xls-tests']['passed_tests'] == 8
    selected = set(json.loads((ROOT / 'package-selection.json').read_text())['selected_packages'])
    strict_args = receipts['strict-44-packages']['argv']
    strict_packages = {strict_args[i + 1] for i, arg in enumerate(strict_args) if arg == '-p'}
    assert len(selected) == 45
    assert strict_packages == selected - {'litchi-odf-common'}
    assert '-A' not in strict_args and '--no-deps' in strict_args
    for name in ['facade-doc', 'non-iwork-rustdoc']:
        assert receipts[name]['environment']['RUSTDOCFLAGS'] == '-D warnings'
    assert b'OK: 9 performance claims validated (strict)' in raw_log(receipts['claims'])
    print(json.dumps({
        'status': 'pass', 'command_receipts': len(receipts),
        'composite_45_package_passed_tests': total,
        'ignored_tests': 98,
        'basis': 'full run plus three complete integration-target reruns; only those test files changed',
        'strict_full_gate': 'open; see ODF, facade, harness and native-resave debt',
        'performance_claim': None,
    }, indent=2))


if __name__ == '__main__':
    main()
