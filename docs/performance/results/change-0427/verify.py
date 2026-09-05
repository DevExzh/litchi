#!/usr/bin/env python3
"""Replay the complete retention bundle without original source or executables."""
import argparse
import gzip
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def artifact(root, name):
    path = root / name
    assert path.resolve().is_relative_to(root.resolve()), name
    return path.read_bytes() if path.exists() else gzip.decompress(path.with_suffix(path.suffix + '.gz').read_bytes())


def verify(root):
    protocol = json.loads((root / 'protocol.json').read_text())
    build = json.loads((root / 'build.json').read_text())
    assert sha((root / 'protocol.json').read_bytes()) == build['protocol_sha256']
    assert sha((root / 'machine.json').read_bytes()) == build['machine_sha256']
    assert sha((root / 'capture.py').read_bytes()) == build['capture_driver_sha256']
    assert sha((root / 'verify-report.py').read_bytes()) == build['verifier_sha256']
    assert sha((root / 'probe-report.py').read_bytes()) == build['probe_sha256']
    assert sha((root / 'summarize.py').read_bytes()) == build['summary_driver_sha256']
    assert sha((root / build['build_receipt']).read_bytes()) == build['build_receipt_sha256']
    checks = json.loads((root / 'expected-checks.json').read_text())
    for name, status in checks.items():
        row = json.loads((root / 'checks' / (name + '.json')).read_text())
        assert row['status'] == status and row['source_unchanged'], name
        assert row['source_before'] == row['source_after'], name
        source = row['source_after']
        source_raw = (root / source['path']).read_bytes()
        assert sha(source_raw) == source['sha256']
        assert len(json.loads(source_raw)) == source['files']
        log = artifact(root, row['log']['path'])
        assert sha(log) == row['log']['sha256'] and len(log) == row['log']['bytes']
    actual = {p.stem for p in (root / 'checks').glob('*.json')
              if 'source_before' in json.loads(p.read_text())}
    assert actual == set(checks)

    index = json.loads((root / 'capture-index.json').read_text())
    assert len(index) == protocol['expected_processes']
    expected_order = [(corpus, row['api'], row['repeat']) for corpus in protocol['corpora']
                      for row in protocol['order_per_corpus']]
    identities = {}
    outputs = {}
    sample_count = 0
    for receipt_name, expected in zip(index, expected_order, strict=True):
        row = json.loads((root / receipt_name).read_text())
        assert (row['corpus'], row['api'], row['repeat']) == expected
        assert row['status'] == 'pass' and row['exit_code'] == 0
        assert row['revision'] == build['revision']
        assert row['binary_sha256'] == build['binary_sha256']
        assert row['protocol_sha256'] == build['protocol_sha256']
        for name, custody in row['artifacts'].items():
            raw = artifact(root, name)
            assert sha(raw) == custody['sha256'] and len(raw) == custody['bytes'], name
        report_path = root / 'capture' / (row['name'] + '.json')
        subprocess.check_output([sys.executable, '-B', str(root / 'verify-report.py'), str(report_path)])
        report = json.loads(report_path.read_text())
        assert report['source_revision'] == build['revision']
        assert report['binary_sha256'] == build['binary_sha256']
        assert report['api'] == row['api'] and report['corpus'] == row['corpus']
        assert report['samples'] == protocol['samples_per_process']
        assert report['warmup'] == protocol['warmups_per_process']
        sample_count += len(report['samples_raw'])
        assert report['binary_bytes'] == build['binary_bytes']
        identity = (report['source_archive_sha256'], report['source_archive_bytes'],
                    report['destination_archive_sha256'], report['destination_archive_bytes'])
        previous = identities.setdefault(row['corpus'], identity)
        assert identity == previous, 'cross-API corpus bytes differ'
        output = (report['expected_output_sha256'], report['expected_output_bytes'])
        assert outputs.setdefault((row['api'], row['corpus']), output) == output, 'same-role repeat output differs'
        subprocess.check_output([sys.executable, '-B', str(root / 'probe-report.py'), str(report_path)])
    assert sample_count == protocol['expected_retained_samples']
    subprocess.check_output([sys.executable, '-B', str(root / 'summarize.py'), '--check'])

    for name, row in json.loads((root / 'compression.json').read_text()).items():
        stored = (root / name).read_bytes()
        raw = gzip.decompress(stored)
        assert sha(stored) == row['stored_sha256'] and len(stored) == row['stored_bytes']
        assert sha(raw) == row['original_sha256'] and len(raw) == row['original_bytes']
    inventory = {}
    for line in (root / 'SHA256SUMS').read_text().splitlines():
        digest, name = line.split('  ', 1)
        assert name not in inventory
        inventory[name] = digest
        assert sha((root / name).read_bytes()) == digest, name
    assert set(inventory) == {str(p.relative_to(root)) for p in root.rglob('*')
                              if p.is_file() and p.name != 'SHA256SUMS'}
    return {'status': 'pass', 'command_receipts': len(checks), 'processes': len(index),
            'samples': sample_count, 'inventory_files': len(inventory), 'performance_claim': None}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--portable', action='store_true')
    args = parser.parse_args()
    result = verify(ROOT)
    if args.portable:
        with tempfile.TemporaryDirectory(prefix='litchi-0427-replay-') as temporary:
            exported = Path(temporary) / 'bundle'
            shutil.copytree(ROOT, exported)
            replay = json.loads(subprocess.check_output([sys.executable, '-B', str(exported / 'verify.py')]))
            assert replay == result
            repeat = exported / 'capture/plain-owned-r2.json'
            receipt_path = exported / 'capture/plain-owned-r2-receipt.json'
            original_report = repeat.read_bytes()
            original_receipt = receipt_path.read_bytes()
            changed = json.loads(original_report)
            changed['expected_output_sha256'] = '0' * 64 if changed['expected_output_sha256'] != '0' * 64 else '1' * 64
            repeat.write_text(json.dumps(changed, indent=2) + '\n')
            custody = json.loads(original_receipt)
            custody['artifacts']['capture/plain-owned-r2.json'] = {
                'sha256': sha(repeat.read_bytes()), 'bytes': repeat.stat().st_size}
            receipt_path.write_text(json.dumps(custody, indent=2) + '\n')
            rejected = subprocess.run([sys.executable, '-B', str(exported / 'verify.py')], capture_output=True)
            assert rejected.returncode != 0 and b'same-role repeat output differs' in rejected.stderr
            repeat.write_bytes(original_report)
            receipt_path.write_bytes(original_receipt)
            validator = exported / 'verify-report.py'
            validator.write_text(validator.read_text() + '\n# custody mutation\n')
            rejected = subprocess.run([sys.executable, '-B', str(exported / 'verify.py')], capture_output=True)
            assert rejected.returncode != 0, 'modified pinned validator was accepted'
        result['portable_export'] = 'pass; copied executable not needed'
        result['pinned_validator_mutation'] = 'rejected'
        result['valid_digest_repeat_output_mutation'] = 'rejected'
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()
