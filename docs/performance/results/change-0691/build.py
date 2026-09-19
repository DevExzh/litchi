#!/usr/bin/env python3
"""Build native and allocation companions serially and bind their source inputs."""
import hashlib
import json
import os
import shutil
import subprocess
import time
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0691'
BIN = ROOT.parent / 'litchi-0691-bin'
BIN.mkdir(exist_ok=True)
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
base = json.loads((P / 'baseline.json').read_text())
for name, digest in base['source_sha256'].items():
    assert sha(ROOT / name) == digest, name
records = []
for label, features in [('native', []), ('allocations', ['--features', 'allocations'])]:
    command = ['cargo', 'build', '--release', '--locked', '--manifest-path', str(P / 'probe/Cargo.toml'),
               '--target-dir', str(TARGET), '-j', '2', *features]
    start = time.monotonic()
    with (P / f'build-{label}.log').open('w') as log:
        result = subprocess.run(command, cwd=ROOT, env={**os.environ, 'RUSTFLAGS': '-D warnings'},
                                stdout=log, stderr=subprocess.STDOUT)
    assert result.returncode == 0, label
    output = BIN / label
    shutil.copy2(TARGET / 'release/probe0691', output)
    records.append(dict(label=label, command=command, exit_code=result.returncode,
                        seconds=time.monotonic()-start, binary=str(output), binary_sha256=sha(output),
                        probe_sha256={str(f.relative_to(P)): sha(f) for f in (P / 'probe').rglob('*') if f.is_file()},
                        source_sha256=base['source_sha256'], rustflags='-D warnings'))
    (P / 'builds.json').write_text(json.dumps(records, indent=2)+'\n')
    print(label, 'built', flush=True)
