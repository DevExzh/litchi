#!/usr/bin/env python3
"""Retain focused XML audit tests with exact source custody."""
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from build import ROOT, TARGET, census, sha

P = Path(__file__).resolve().parent
label = sys.argv[1]
assert label in ['baseline', 'candidate', 'restored']
command = ['cargo','test','-p','xml-minifier','--all-features','--locked','--','--test-threads=1']
env = dict(os.environ, CARGO_BUILD_JOBS='2', CARGO_TARGET_DIR=str(TARGET), RUSTFLAGS='-D warnings')
before = census()
start = time.monotonic()
log_path = P/f'focused-{label}.log'
with log_path.open('w') as log:
    result = subprocess.run(command, cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT)
assert census() == before
(P/f'focused-{label}.json').write_text(json.dumps(dict(command=command, exit_code=result.returncode,
    seconds=time.monotonic()-start, source=before, script_sha256=sha(Path(__file__)),
    log_sha256=sha(log_path), environment={k:env.get(k) for k in ['CARGO_BUILD_JOBS','CARGO_TARGET_DIR','RUSTFLAGS']}),indent=2)+'\n')
print(label, 'focused tests exit', result.returncode, flush=True)
raise SystemExit(result.returncode)
