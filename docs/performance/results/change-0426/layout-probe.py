#!/usr/bin/env python3
"""Capture static XLS type sizes from an explicitly selected compiled library."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--library', type=Path, required=True)
    args = parser.parse_args()
    library = args.library.resolve()
    log = ROOT / 'checks/layout-after.log'
    receipt = ROOT / 'checks/layout-after.json'
    assert not receipt.exists() and not log.exists()
    with tempfile.TemporaryDirectory(prefix='litchi-0426-layout-') as temporary:
        binary = Path(temporary) / 'layout'
        argv = ['rustup', 'run', '1.98.1', 'rustc', '--edition', '2024',
                str(ROOT / 'layout.rs'), '-L', 'dependency=' + str(library.parent),
                '--extern', 'litchi_xls=' + str(library), '-o', str(binary)]
        record = {
            'change': 426, 'argv': argv,
            'compiled_rlib': str(library.relative_to(REPO)),
            'compiled_rlib_sha256': sha(library),
            'source_model_sha256': sha(REPO / 'crates/litchi-xls/src/formula_metadata/model.rs'),
            'probe_source_sha256': sha(ROOT / 'layout.rs'),
            'scope': 'Rust static type layout; not operation allocations, RSS or performance',
        }
        subprocess.run(argv, cwd=REPO, check=True)
        result = subprocess.run([str(binary)], capture_output=True, check=True)
        log.write_bytes(result.stdout)
        record.update(exit_code=result.returncode,
                      finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                      sizes=json.loads(result.stdout), binary_sha256=sha(binary))
        receipt.write_text(json.dumps(record, indent=2) + '\n')
        print(json.dumps(record['sizes']))


if __name__ == '__main__':
    main()
