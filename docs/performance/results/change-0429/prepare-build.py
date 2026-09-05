#!/usr/bin/env python3
"""Bind a completed release build and copy its executable for isolated capture."""
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()



def verify_current_sources(manifest):
    spec = importlib.util.spec_from_file_location('custody', ROOT / 'check.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    assert module.sources() == manifest, 'tracked or untracked source manifest changed'


def main():
    receipt_path = ROOT / 'checks/release-build.json'
    receipt = json.loads(receipt_path.read_text())
    assert receipt['status'] == 'pass' and receipt['source_unchanged']
    assert receipt['source_before'] == receipt['source_after']
    verify_current_sources(receipt['source_after'])
    revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO).decode().strip()
    assert receipt['revision'] == revision
    subprocess.run(['git', 'diff', '--exit-code', revision, '--', '*.rs', '*.toml', '*.lock'], cwd=REPO, check=True)
    original = REPO / 'tools/perf-baseline/target/release/litchi-perf-baseline'
    temporary = Path('/tmp/litchi-goal-0429-binaries')
    temporary.mkdir(exist_ok=False)
    binary = temporary / original.name
    shutil.copy2(original, binary)
    assert sha(original) == sha(binary)
    target = ROOT / 'build.json'
    assert not target.exists()
    target.write_text(json.dumps({
        'revision': revision, 'baseline_revision': '995e217ba5624d6e9926a7bf2677191067f2bd6c',
        'build_receipt': str(receipt_path.relative_to(ROOT)), 'build_receipt_sha256': sha(receipt_path),
        'source_manifest': receipt['source_after'],
        'original_binary': str(original.relative_to(REPO)), 'capture_binary': str(binary),
        'binary_sha256': sha(binary), 'binary_bytes': binary.stat().st_size,
        'protocol_sha256': sha(ROOT / 'protocol.json'),
        'planned_checks_sha256': sha(ROOT / 'planned-checks.json'),
        'machine_sha256': sha(ROOT / 'machine.json'),
        'capture_driver_sha256': sha(ROOT / 'capture.py'),
        'verifier_sha256': sha(ROOT / 'verify-report.py'),
        'replay_verifier_sha256': sha(ROOT / 'verify.py'),
        'probe_sha256': sha(ROOT / 'probe-report.py'),
        'summary_driver_sha256': sha(ROOT / 'summarize.py'),
        'strict_comparison_driver_sha256': sha(ROOT / 'compare-strict.py'),
        'additional_artifact_sha256': {name: sha(ROOT / name) for name in ['check.py', 'profiles.py', 'profile-audit.py', 'native-oracles.py', 'native-image-oracles.json', 'native-inputs/original.pptx', 'native-inputs/libreoffice.pptx', 'native-inputs/poi-slide.pptx', 'native-inputs/poi-video.pptx', 'shapes-static-oracles.json']},
        'rust_toolchain': '1.98.1', 'profile': 'release',
        'scope': 'one frozen provider baseline build with ZIP short-read correctness enabler; no causal before/after speedup',
    }, indent=2) + '\n')
    print(str(binary))


if __name__ == '__main__':
    main()
