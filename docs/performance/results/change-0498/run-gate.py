#!/usr/bin/env python3
"""Run one recorded 0498 gate with pinned tooling and retained output."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
commands = json.loads((HERE / 'gate-commands.json').read_text())
name = sys.argv[1]
command = ['taskset', '-c', '16-31', *commands[name]]
env = os.environ.copy()
env.update(RUSTUP_TOOLCHAIN='1.98.1', CARGO_INCREMENTAL='0',
           CARGO_PROFILE_RELEASE_DEBUG='0', CARGO_BUILD_JOBS='2',
           CARGO_TARGET_DIR='/home/zhuhe/.cache/litchi-goal-0498/target',
           RUSTDOCFLAGS='-D warnings')
started = time.time()
log = HERE / 'checks' / (name + '.log')
if log.exists():
    attempt = 1
    while log.with_name(f'{name}.attempt{attempt}.log').exists():
        attempt += 1
    log.rename(log.with_name(f'{name}.attempt{attempt}.log'))
    old_receipt = HERE / 'checks' / (name + '.json')
    if old_receipt.exists():
        old_receipt.rename(old_receipt.with_name(f'{name}.attempt{attempt}.json'))
with log.open('wb') as output:
    result = subprocess.run(command, cwd=ROOT, env=env, stdout=output,
                            stderr=subprocess.STDOUT)
receipt = dict(command=command, exit_code=result.returncode,
               elapsed_seconds=time.time()-started,
               log_sha256=hashlib.sha256(log.read_bytes()).hexdigest())
(HERE / 'checks' / (name + '.json')).write_text(json.dumps(receipt, indent=2)+'\n')
print(json.dumps(receipt))
print(log.read_text(errors='replace')[-3500:])
sys.exit(result.returncode)
