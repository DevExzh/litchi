#!/usr/bin/env python3
"""Bind copied executables to the completed build and frozen capture contract."""
import argparse
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--role', choices=('before', 'after'), required=True)
    args = parser.parse_args()
    name = f'checks/{args.role}-build.json'
    receipt = json.loads((ROOT / name).read_text())
    assert receipt['status'] == 'pass' and receipt['source_unchanged'] is True
    digest = lambda name: hashlib.sha256((ROOT / name).read_bytes()).hexdigest()
    row = {
        'change': 442, 'role': args.role, 'revision': receipt['revision'],
        'source_manifest': receipt['source_after'],
        'binaries': json.loads((ROOT / args.role / 'binary-copies.json').read_text()),
        'protocol_sha256': digest('protocol.json'),
        'capture_driver_sha256': digest('capture.py'),
        'profile_driver_sha256': digest('profile.py'), 'build_receipt': name,
    }
    with (ROOT / args.role / 'build.json').open('x') as stream:
        stream.write(json.dumps(row, indent=2) + '\n')
    print('Bound ' + args.role + ' build')


if __name__ == '__main__':
    main()
