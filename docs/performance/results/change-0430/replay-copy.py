#!/usr/bin/env python3
"""Verify a sealed copy while retaining a command receipt in the original bundle."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tag', required=True)
    args = parser.parse_args()
    assert re.fullmatch(r'[a-z0-9_-]+', args.tag)
    receipt = ROOT / (args.tag + '-export.json')
    assert not receipt.exists()
    row = {
        'status': 'running',
        'started_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
        'input_inventory_sha256': digest(ROOT / 'SHA256SUMS'),
        'input_verifier_sha256': digest(ROOT / 'verify.py'),
        'driver_sha256': digest(Path(__file__)),
        'tag': args.tag,
    }
    result = None
    try:
        with tempfile.TemporaryDirectory(prefix='litchi-0430-replay-') as temporary:
            copied = Path(temporary) / 'bundle'
            # Copy before check.py creates its running receipt and log. The
            # verifier consequently sees exactly the sealed inventory, with
            # no exception for a concurrently changing command receipt.
            shutil.copytree(ROOT, copied)
            assert digest(copied / 'SHA256SUMS') == row['input_inventory_sha256']
            assert digest(copied / 'verify.py') == row['input_verifier_sha256']
            row['temporary_directory'] = temporary
            result = subprocess.run([
                sys.executable, '-B', str(ROOT / 'check.py'), '--tag', args.tag,
                '--', sys.executable, '-B', str(copied / 'verify.py'), '--portable-check',
            ], check=False)
            row['exit_code'] = result.returncode
    finally:
        row['temporary_directory_absent'] = (
            'temporary_directory' in row and not Path(row['temporary_directory']).exists())
        row['status'] = ('pass' if result is not None and result.returncode == 0
                         and row['temporary_directory_absent'] else 'failed')
        row['finished_utc'] = datetime.datetime.now(datetime.timezone.utc).isoformat()
        receipt.write_text(json.dumps(row, indent=2) + '\n')
    print(json.dumps(row))
    return 0 if row['status'] == 'pass' else 1


if __name__ == '__main__':
    raise SystemExit(main())
