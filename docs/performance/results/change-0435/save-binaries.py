#!/usr/bin/env python3
"""Retain a release build's executables before a later build can replace them."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--role', choices=('before', 'after'), required=True)
    args = parser.parse_args()
    receipt = json.loads((ROOT / 'checks' / f'{args.role}-build.json').read_text())
    assert receipt['status'] == 'pass' and receipt['source_unchanged'] is True
    directory = Path('/tmp/litchi-goal-0435-binaries') / args.role
    directory.mkdir(parents=True, exist_ok=False)
    copies = {}
    for mode, name in [('normal', 'litchi-perf-baseline'),
                       ('allocator', 'litchi-perf-baseline-alloc')]:
        source = REPO / 'tools/perf-baseline/target/release' / name
        destination = directory / mode
        shutil.copy2(source, destination)
        assert sha(source) == sha(destination)
        copies[mode] = {'path': str(destination), 'bytes': destination.stat().st_size,
                        'sha256': sha(destination)}
    target = ROOT / args.role / 'binary-copies.json'
    target.parent.mkdir(exist_ok=True)
    with target.open('x') as stream:
        json.dump(copies, stream, indent=2)
        stream.write('\n')
    print(json.dumps(copies))


if __name__ == '__main__':
    main()
