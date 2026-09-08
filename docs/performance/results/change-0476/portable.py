#!/usr/bin/env python3
"""Replay a sealed fresh copy and reject resealed mutations without runtime trees."""
from pathlib import Path
import datetime
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def seal(root):
    rows = []
    for path in sorted(root.rglob('*')):
        assert not path.is_symlink()
        if path.is_file() and path != root / 'SHA256SUMS':
            rows.append(f'{sha(path)}  {path.relative_to(root).as_posix()}\n')
    (root / 'SHA256SUMS').write_text(''.join(rows))


def main():
    assert not (ROOT / 'portable.json').exists()
    record = dict(started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        driver_sha256=sha(Path(__file__)), input_seal_sha256=sha(ROOT / 'SHA256SUMS'))
    with tempfile.TemporaryDirectory(prefix='litchi-goal-0476-portable-') as directory:
        bundle = Path(directory) / 'bundle'
        shutil.copytree(ROOT, bundle)
        argv = [sys.executable, '-B', str(bundle / 'verify.py')]
        baseline = subprocess.run(argv, cwd=bundle, capture_output=True, text=True)
        record['baseline'] = dict(exit_code=baseline.returncode, stdout=baseline.stdout, stderr=baseline.stderr)
        assert baseline.returncode == 0, baseline.stderr
        record['mutations'] = {}
        cases = [
            ('derived-summary', 'summary.json'),
            ('receipt-command-omitted', 'captures/R1-control-normal-tiny/receipt.json'),
            ('producer-output-omitted', 'captures/R1-control-normal-tiny/report.json'),
            ('source-manifest-omitted', 'candidate-source.json'),
            ('guard-shape-duplicated', 'captures/G1-control/report.json'),
        ]
        for label, relative in cases:
            path = bundle / relative
            backups = {path: path.read_bytes()}
            value = json.loads(backups[path])
            if label == 'derived-summary':
                value['acceptance']['R1_percent'] += 1
            elif label == 'receipt-command-omitted':
                del value['argv']
            elif label == 'producer-output-omitted':
                del value['results'][0]['output_sha256']
            elif label == 'source-manifest-omitted':
                del value['source_manifest']
            else:
                rows = value['results']
                first = rows[0]
                other = next(i for i, row in enumerate(rows)
                    if row['case'] == first['case'] and row['corpus']['shape'] != first['corpus']['shape'])
                rows[other] = json.loads(json.dumps(first))
            path.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')
            if path.name == 'report.json':
                receipt_path = path.with_name('receipt.json')
                backups[receipt_path] = receipt_path.read_bytes()
                receipt = json.loads(backups[receipt_path])
                receipt['artifacts']['report.json'] = dict(bytes=path.stat().st_size, sha256=sha(path))
                receipt_path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n')
            seal(bundle)
            result = subprocess.run(argv, cwd=bundle, capture_output=True, text=True)
            record['mutations'][label] = dict(exit_code=result.returncode, stdout=result.stdout, stderr=result.stderr)
            assert result.returncode != 0, f'mutation accepted: {label}'
            for original, data in backups.items():
                original.write_bytes(data)
            seal(bundle)
        replay = subprocess.run(argv, cwd=bundle, capture_output=True, text=True)
        assert replay.returncode == 0, replay.stderr
        record['restored_replay'] = dict(exit_code=replay.returncode, stdout=replay.stdout, stderr=replay.stderr)
    record.update(finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        temporary_copy_removed=not Path(directory).exists())
    with (ROOT / 'portable.json').open('x') as stream:
        json.dump(record, stream, indent=2, sort_keys=True)
        stream.write('\n')
    print('portable copy replay and five resealed mutation rejections passed')


if __name__ == '__main__':
    main()
