#!/usr/bin/env python3
"""Build the existing native harness against an exact source census."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0705'

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def census():
    return {str(p.relative_to(ROOT)): sha(p)
            for folder in ['crates', 'tools/perf-baseline']
            for p in sorted((ROOT / folder).rglob('*'))
            if p.is_file() and 'target' not in p.parts
            and (p.suffix == '.rs' or p.name in ['Cargo.toml', 'Cargo.lock'])}

def main():
    before = census()
    workspace = [ROOT / 'Cargo.toml', ROOT / 'Cargo.lock', *(ROOT / '.cargo').rglob('*')]
    inputs = {str(p.relative_to(ROOT)): sha(p) for p in workspace if p.is_file()}
    (P / 'workspace-inputs.json').write_text(json.dumps(inputs, indent=2) + '\n')
    constraints = json.loads((P.parent / 'change-0704/baseline.json').read_text())['constraints_sha256']
    assert all(sha(ROOT / p) == h for p, h in constraints.items())
    (P / 'source-manifest.json').write_text(json.dumps(before, indent=2) + '\n')
    (P / 'constraints.json').write_text(json.dumps(constraints, indent=2) + '\n')
    commands = [['rustc', '-Vv'], ['cargo', '-V'], ['uname', '-a'], ['lscpu'], ['free', '-b']]
    host = []
    for command in commands:
        r = subprocess.run(command, capture_output=True, text=True)
        host.append(dict(command=command, exit_code=r.returncode, stdout=r.stdout, stderr=r.stderr))
    (P / 'host.json').write_text(json.dumps(host, indent=2) + '\n')
    command = ['cargo', 'build', '--release', '--locked', '--manifest-path',
               'tools/perf-baseline/Cargo.toml', '--bin', 'litchi-perf-baseline',
               '--target-dir', str(TARGET), '-j', '2']
    start = time.monotonic()
    with (P / 'build.log').open('w') as log:
        r = subprocess.run(command, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT)
    assert census() == before, 'source changed during build'
    assert all(sha(ROOT / p) == h for p, h in inputs.items())
    binary = TARGET / 'release/litchi-perf-baseline'
    receipt = dict(command=command, exit_code=r.returncode, seconds=time.monotonic()-start,
                   revision=subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip(),
                   binary=str(binary), binary_sha256=sha(binary) if r.returncode == 0 else None,
                   source_manifest_sha256=sha(P / 'source-manifest.json'),
                   environment={k: os.environ.get(k) for k in ['RUSTFLAGS', 'LD_PRELOAD', 'MALLOC_CONF', 'GLIBC_TUNABLES']})
    (P / 'build.json').write_text(json.dumps(receipt, indent=2) + '\n')
    assert r.returncode == 0

if __name__ == '__main__':
    main()
