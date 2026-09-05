#!/usr/bin/env python3
"""Replay retained command, source, log, census and layout custody without Cargo."""
import gzip
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parent


def digest(raw):
    return hashlib.sha256(raw).hexdigest()


def raw_log(record):
    path = ROOT / record['log']['path']
    raw = path.read_bytes() if path.exists() else gzip.decompress(
        path.with_suffix(path.suffix + '.gz').read_bytes())
    assert len(raw) == record['log']['bytes'], path
    assert digest(raw) == record['log']['sha256'], path
    return raw


def main():
    expected = json.loads((ROOT / 'expected-checks.json').read_text())
    receipts = {}
    logs = {}
    for name, status in expected.items():
        record = json.loads((ROOT / 'checks' / (name + '.json')).read_text())
        assert record['status'] == status, name
        assert record['source_unchanged'], name
        assert record['source_before'] == record['source_after'], name
        assert (record['exit_code'] == 0) == (status == 'pass'), name
        logs[name] = raw_log(record)
        receipts[name] = record
    actual = {p.stem for p in (ROOT / 'checks').glob('*.json')
              if 'source_before' in json.loads(p.read_text())}
    assert actual == set(expected), 'unclassified or missing command receipt'

    outcomes = json.loads((ROOT / 'outcomes.json').read_text())
    final_sources = receipts['xls-tests-accepted']['source_after']
    for name, counts in outcomes['test_counts'].items():
        rows = re.findall(
            rb'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',
            logs[name])
        observed = [sum(int(row[i]) for row in rows) for i in range(3)]
        assert observed == counts, (name, observed, counts)
    for name in outcomes['final_source_checks']:
        assert receipts[name]['source_after'] == final_sources, name
        assert receipts[name]['status'] == ('failed' if name == 'facade-strict' else 'pass'), name
    for name in ['xls-rustdoc', 'facade-rustdoc']:
        assert receipts[name]['environment']['RUSTDOCFLAGS'] == '-D warnings'
    assert b'OK: 9 performance claims validated (strict)' in logs['claims']
    assert '-A' not in receipts['xls-strict-verified']['argv']
    debt = json.loads((ROOT / 'checks/facade-strict-debt.json').read_text())
    assert debt['diagnostic_count'] == 18 and debt['same_message_and_source_file_multiset']
    assert digest(logs['facade-strict']) == debt['current_log_sha256']

    census = json.loads((ROOT / 'checks/formula-census.json').read_text())
    assert census['formula_count'] == 1416
    assert len(census['records_with_extra']) == 5
    for row in census['records_with_extra']:
        tokens = bytes.fromhex(row['tokens_hex'])
        extra = bytes.fromhex(row['extra_hex'])
        assert len(tokens) == row['token_bytes']
        assert 22 + len(tokens) + len(extra) == row['payload_bytes']
        assert len(extra) == 10 and extra[:2] == b'\x01\x00'
        values = [int.from_bytes(extra[i:i + 2], 'little') for i in range(2, 10, 2)]
        assert values[0] <= values[1] and values[2] <= values[3] <= 255
    fixture = ROOT.parents[3] / census['fixture']
    if fixture.is_file():
        assert digest(fixture.read_bytes()) == census['fixture_sha256']
        replay = subprocess.check_output([sys.executable, '-B', str(ROOT / 'formula-census.py')])
        assert json.loads(replay) == census

    for name, row in json.loads((ROOT / 'compressed-logs.json').read_text()).items():
        stored = (ROOT / name).read_bytes()
        assert len(stored) == row['stored_bytes']
        assert digest(stored) == row['stored_sha256']
        raw = gzip.decompress(stored)
        assert len(raw) == row['original_bytes']
        assert digest(raw) == row['original_sha256']

    before = json.loads((ROOT / 'checks/layout-before.json').read_text())
    after = json.loads((ROOT / 'checks/layout-after.json').read_text())
    assert before['exit_code'] == after['exit_code'] == 0
    assert before['sizes'] == outcomes['layout_before']
    assert after['sizes'] == outcomes['layout_after']
    assert after['source_model_sha256'] == final_sources[
        'crates/litchi-xls/src/formula_metadata/model.rs']
    # The initial probe retained observed sizes in its receipt and compiler
    # diagnostics separately; it did not retain a second stdout log.
    assert gzip.decompress((ROOT / 'checks/layout-before-build.log.gz').read_bytes()) == b''
    assert json.loads(gzip.decompress((ROOT / 'checks/layout-after.log.gz').read_bytes())) == after['sizes']

    inventory = {}
    for line in (ROOT / 'SHA256SUMS').read_text().splitlines():
        sha, name = line.split('  ', 1)
        assert name not in inventory
        inventory[name] = sha
        assert digest((ROOT / name).read_bytes()) == sha, name
    files = {str(p.relative_to(ROOT)) for p in ROOT.rglob('*') if p.is_file()}
    assert set(inventory) == files - {'SHA256SUMS'}
    print(json.dumps({
        'status': 'pass', 'command_receipts': len(receipts),
        'inventory_files': len(inventory), 'test_counts': outcomes['test_counts'],
        'facade_correctness': 'six baseline failures closed',
        'strict_full_gate': 'open; scoped XLS gate passes',
        'performance_claim': None,
    }, indent=2))


if __name__ == '__main__':
    main()
