#!/usr/bin/env python3
"""Run new focused tests against baseline or candidate with source bindings."""
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
phase = sys.argv[1]
assert phase in ['baseline', 'candidate', 'retained']
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
files = ['crates/litchi-ooxml-common/src/mce/codec.rs',
         'crates/litchi-ooxml-common/src/mce/tests.rs']
source = {name: sha(ROOT / name) for name in files}
if phase == 'baseline':
    base = json.loads((P / 'baseline.json').read_text())
    assert source[files[0]] == base['source_sha256'][files[0]]
elif phase == 'retained':
    rejection = json.loads((P / 'rejection.json').read_text())
    assert source == {name: rejection['final_source_sha256'][name] for name in files}
command = ['cargo', 'test', '-p', 'litchi-ooxml-common', '--lib', '--locked',
           'mce::', '--', '--test-threads=1']
env = dict(os.environ, CARGO_TARGET_DIR=str(ROOT.parent / 'litchi-target-0698'),
           CARGO_BUILD_JOBS='2', RUSTFLAGS='-D warnings')
log = P / {'baseline': 'tests-focused-baseline.log', 'candidate': 'tests-focused.log',
           'retained': 'tests-focused-retained.log'}[phase]
with log.open('w') as output:
    result = subprocess.run(command, cwd=ROOT, env=env, stdout=output, stderr=subprocess.STDOUT)
assert {name: sha(ROOT / name) for name in files} == source
(P / ('focused-' + phase + '.json')).write_text(json.dumps(dict(
    phase=phase, command=command, source_sha256=source, exit_code=result.returncode,
    log=log.name, log_sha256=sha(log), rustflags='-D warnings',
), indent=2) + '\n')
print(phase, 'focused tests exit', result.returncode)
raise SystemExit(result.returncode)
