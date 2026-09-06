#!/usr/bin/env python3
"""Replay a sealed bundle, retaining the terminal receipt after replay finishes."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent
DRIVERS = ('replay.py', 'compare.py', 'verify.py', 'check.py', 'capture.py',
           'verify-report.py', 'seal.py', 'protocol.json')


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tag', required=True)
    args = parser.parse_args()
    assert re.fullmatch(r'[a-z0-9_-]+', args.tag)
    receipt = ROOT / 'checks' / (args.tag + '.json')
    log = receipt.with_suffix('.log')
    assert not receipt.exists() and not log.exists() and not log.with_suffix('.log.gz').exists()
    inventory = ROOT / 'SHA256SUMS'
    assert inventory.is_file(), 'seal the complete bundle before replay'
    drivers = {name: digest(ROOT / name) for name in DRIVERS}
    inventory_before = digest(inventory)
    command = [sys.executable, '-B', str(ROOT / 'verify.py'), '--portable-check']
    started = now()
    # Keep the running log outside the sealed bundle. No running receipt can
    # be mistaken for a completed validation, and inventory stays immutable.
    with tempfile.TemporaryDirectory(prefix='litchi-0431-replay-') as temporary:
        temporary_log = Path(temporary) / 'replay.log'
        with temporary_log.open('xb') as stream:
            result = subprocess.run(command, stdout=stream, stderr=subprocess.STDOUT)
        raw = temporary_log.read_bytes()
    drivers_unchanged = drivers == {name: digest(ROOT / name) for name in DRIVERS}
    inventory_unchanged = inventory_before == digest(inventory)
    passed = result.returncode == 0 and drivers_unchanged and inventory_unchanged
    log.write_bytes(raw)
    record = {
        'change': 431, 'status': 'pass' if passed else 'failed',
        'command': command, 'exit_code': result.returncode,
        'started_utc': started, 'finished_utc': now(),
        'driver_hashes': drivers, 'drivers_unchanged': drivers_unchanged,
        'replayed_inventory_sha256': inventory_before,
        'inventory_unchanged': inventory_unchanged,
        'log': {'path': str(log.relative_to(ROOT)), 'bytes': len(raw),
                'sha256': hashlib.sha256(raw).hexdigest()},
        'scope': 'Terminal external replay receipt; reseal after adding this receipt and log.',
    }
    receipt.write_text(json.dumps(record, indent=2) + '\n')
    print(json.dumps({'status': record['status'], 'exit_code': result.returncode}))
    return 0 if passed else 1


if __name__ == '__main__':
    sys.exit(main())
