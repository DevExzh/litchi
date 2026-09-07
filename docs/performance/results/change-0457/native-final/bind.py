#!/usr/bin/env python3
"""Retain the native append probe and bind its completed build before capture."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[4]


def sha(path):
    digest = hashlib.sha256()
    with path.open('rb') as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--build-receipt', type=Path, required=True)
    parser.add_argument('--binary', type=Path, required=True)
    parser.add_argument('--retained-binary', type=Path, required=True)
    args = parser.parse_args()
    binding = ROOT / 'binary-binding.json'
    assert not binding.exists() and not (ROOT / 'runs').exists()
    assert not args.retained_binary.exists()
    build = json.loads(args.build_receipt.read_text())
    assert build['status'] == 'pass' and build['exit_code'] == 0
    assert build['source_unchanged'] and build['source_before'] == build['source_after']
    assert build['cwd'] == str(REPO)
    assert 'build' in build['argv'] and 'odp_native_tail_append_probe' in build['argv']
    assert build['driver_sha256'] == sha(ROOT.parent / 'check.py')
    manifest = ROOT.parent / build['source_after']['path']
    assert sha(manifest) == build['source_after']['sha256']
    assert len(json.loads(manifest.read_text())) == build['source_after']['files']
    assert args.binary.is_file() and not args.binary.is_symlink()
    assert os.access(args.binary, os.X_OK)
    args.retained_binary.parent.mkdir(parents=True, exist_ok=True)
    with args.binary.open('rb') as source, args.retained_binary.open('xb') as target:
        shutil.copyfileobj(source, target)
    args.retained_binary.chmod(args.binary.stat().st_mode & 0o777)
    assert sha(args.binary) == sha(args.retained_binary)
    record = {
        'change': 457,
        'binary': {'path': str(args.retained_binary.resolve()),
                   'bytes': args.retained_binary.stat().st_size,
                   'sha256': sha(args.retained_binary)},
        'build_receipt': {'path': os.path.relpath(args.build_receipt.resolve(), ROOT),
                          'sha256': sha(args.build_receipt)},
        'source_manifest': {'path': os.path.relpath(manifest, ROOT),
                            'sha256': sha(manifest), 'files': build['source_after']['files']},
        'inventory_sha256': sha(ROOT / 'staticinventory.json'),
        'oracle_sha256': sha(ROOT / 'verify-output.py'),
        'runner_sha256': sha(ROOT / 'run.py'),
        'binder_sha256': sha(Path(__file__)),
    }
    with binding.open('x') as target:
        json.dump(record, target, indent=2)
        target.write('\n')
    print(json.dumps({'status': 'pass', 'binding_sha256': sha(binding)}))


if __name__ == '__main__':
    main()
