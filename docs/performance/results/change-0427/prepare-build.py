#!/usr/bin/env python3
"""Bind a completed release build and copy its executable for isolated capture."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    receipt_path = ROOT / 'checks/release-build.json'
    receipt = json.loads(receipt_path.read_text())
    assert receipt['status'] == 'pass' and receipt['source_unchanged']
    assert receipt['source_before'] == receipt['source_after']
    revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO).decode().strip()
    assert receipt['revision'] == revision
    subprocess.run(['git', 'diff', '--exit-code', revision, '--', '*.rs', '*.toml', '*.lock'], cwd=REPO, check=True)
    original = REPO / 'tools/perf-baseline/target/release/litchi-perf-baseline-alloc'
    temporary = Path('/tmp/litchi-goal-0427-binaries')
    temporary.mkdir(exist_ok=False)
    binary = temporary / original.name
    shutil.copy2(original, binary)
    assert sha(original) == sha(binary)
    target = ROOT / 'build.json'
    assert not target.exists()
    target.write_text(json.dumps({
        'revision': revision, 'production_revision': 'd2f98b02e1d620c84359604a34c4ffadd9213cf3',
        'build_receipt': str(receipt_path.relative_to(ROOT)), 'build_receipt_sha256': sha(receipt_path),
        'source_manifest': receipt['source_after'],
        'original_binary': str(original.relative_to(REPO)), 'capture_binary': str(binary),
        'binary_sha256': sha(binary), 'binary_bytes': binary.stat().st_size,
        'protocol_sha256': sha(ROOT / 'protocol.json'),
        'machine_sha256': sha(ROOT / 'machine.json'),
        'capture_driver_sha256': sha(ROOT / 'capture.py'),
        'verifier_sha256': sha(ROOT / 'verify-report.py'),
        'probe_sha256': sha(ROOT / 'probe-report.py'),
        'summary_driver_sha256': sha(ROOT / 'summarize.py'),
        'rust_toolchain': '1.98.1', 'profile': 'release',
        'scope': 'one frozen diagnostic build; no production optimization or before/after speedup',
    }, indent=2) + '\n')
    print(str(binary))


if __name__ == '__main__':
    main()
