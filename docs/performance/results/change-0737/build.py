#!/usr/bin/env python3
"""Serial probe gates, exact source custody, preserved build attempts."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0737'
BIN = ROOT.parent / 'litchi-0737-bin'


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def write(p, value):
    p.write_text(json.dumps(value, indent=2) + '\n')


def census(root):
    return {str(p.relative_to(root)): sha(p) for p in sorted(root.rglob('*')) if p.is_file()}


variant = sys.argv[1]
assert variant in ('legacy', 'controls')
probe = P / ('legacy-probe' if variant == 'legacy' else 'probe')
source = json.loads((P / 'base.json').read_text())['source']
assert all(sha(ROOT / f) == h for f, h in source.items())
assert all(sha(ROOT / f) == h for f, h in json.loads((P / 'constraints.json').read_text()).items())
before = census(probe)
i = 0
while (P / f'build-{i}').exists():
    i += 1
out = P / f'build-{i}'
out.mkdir()
shutil.copytree(probe, out / 'probe')
env = dict(os.environ, CARGO_TARGET_DIR=str(TARGET), CARGO_BUILD_JOBS='2', RUSTDOCFLAGS='-D warnings')
m = str(probe / 'Cargo.toml')
common = ['--manifest-path', m, '--release', '--offline', '--locked']
commands = [['cargo', 'fmt', '--manifest-path', m, '--', '--check'],
            ['cargo', 'test', *common, '--lib'],
            ['cargo', 'clippy', *common, '--all-targets', '--', '-D', 'warnings'],
            ['cargo', 'doc', *common, '--no-deps'],
            ['cargo', 'build', *common, '--bins']]
runs = []
for n, command in enumerate(commands):
    start = time.monotonic()
    with (out / f'{n}.log').open('wb') as handle:
        result = subprocess.run(command, cwd=ROOT, env=env, stdout=handle, stderr=subprocess.STDOUT)
    runs.append(dict(command=command, exit_code=result.returncode, seconds=time.monotonic()-start,
                     output=f'build-{i}/{n}.log', sha256=sha(out / f'{n}.log')))
    write(out / 'manifest.json', dict(variant=variant, source=source, probe=before, runs=runs))
    print(variant, n, result.returncode, flush=True)
    assert result.returncode == 0
assert census(probe) == before
assert all(sha(ROOT / f) == h for f, h in source.items())
BIN.mkdir(exist_ok=True)
binaries = {}
for lane, name in [('native', 'ole_format_save_probe'), ('allocation', 'ole_format_save_probe_alloc')]:
    destination = BIN / f'{variant}-{name}'
    assert not destination.exists()
    shutil.copy2(TARGET / 'release' / name, destination)
    binaries[lane] = dict(path=str(destination), bytes=destination.stat().st_size, sha256=sha(destination))
write(P / f'{variant}-build.json', dict(source=source, probe=before,
       quality=f'build-{i}/manifest.json', quality_sha256=sha(out / 'manifest.json'), binaries=binaries))
print('PASS five exact-source probe gates', variant)
