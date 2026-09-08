#!/usr/bin/env python3
"""Record an exclusive, immutable validation attempt and its tool identities."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def main():
    label, *argv = sys.argv[1:]
    assert label and '/' not in label and argv
    directory = ROOT / 'validation'
    directory.mkdir(exist_ok=True)
    prefix = directory / label
    receipt = dict(argv=argv, cwd=str(ROOT),
        started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        python_sources={p.name: sha(p) for p in sorted(ROOT.glob('*.py'))},
        protocol_sha256=sha(ROOT / 'protocol.json'),
        source_manifest_sha256=sha(ROOT / 'reuse/sources/source.json'),
        environment={'PYTHONDONTWRITEBYTECODE': '1', 'RUSTUP_TOOLCHAIN': '1.98.1'})
    with prefix.with_suffix('.started.json').open('x') as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True); stream.write('\n')
    with prefix.with_suffix('.stdout').open('xb') as out, prefix.with_suffix('.stderr').open('xb') as err:
        result = subprocess.run(argv, cwd=ROOT,
            env=dict(os.environ, PYTHONDONTWRITEBYTECODE='1', RUSTUP_TOOLCHAIN='1.98.1'), stdout=out, stderr=err)
    receipt.update(exit_code=result.returncode,
        finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        artifacts={prefix.with_suffix(s).name: dict(sha256=sha(prefix.with_suffix(s)),
            bytes=prefix.with_suffix(s).stat().st_size) for s in ('.stdout', '.stderr')})
    with prefix.with_suffix('.json').open('x') as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True); stream.write('\n')
    print(label, result.returncode, flush=True)
    raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()
