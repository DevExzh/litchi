#!/usr/bin/env python3
"""Run source-bound crate and harness checks serially before measurements."""
import datetime
import json
import os
import subprocess
import time
from run import HERE, REPO, SCRATCH, sha, sources, write

env = dict(os.environ, CARGO_TARGET_DIR=str(SCRATCH/'target'), CARGO_BUILD_JOBS='2', CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_DEBUG='0', CARGO_PROFILE_DEV_DEBUG='0', CARGO_PROFILE_TEST_DEBUG='0', TMPDIR=str(SCRATCH), RUSTDOCFLAGS='-D warnings')
commands = [
    ('fmt', ['cargo','fmt','--all','--','--check']),
    ('odt-tests', ['cargo','test','--release','--locked','-p','litchi-odt']),
    ('odt-clippy', ['cargo','clippy','--release','--locked','-p','litchi-odt','--all-targets','--','-D','warnings']),
    ('odt-rustdoc', ['cargo','doc','--release','--locked','-p','litchi-odt','--no-deps']),
    ('harness-tests', ['cargo','test','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--lib']),
    ('harness-clippy', ['cargo','clippy','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--lib','--bin','litchi-perf-baseline','--','-D','warnings']),
]
for name, command in commands:
    before = sources()
    assert before == json.loads((HERE / 'source-manifest.json').read_text())
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    tick = time.monotonic()
    with (HERE / f'{name}.log').open('x') as log:
        result = subprocess.run(command,cwd=REPO,env=env,stdout=log,stderr=subprocess.STDOUT)
    unchanged = sources()==before
    write(HERE / f'{name}-receipt.json', {'command':command,'started_utc':started,'elapsed_seconds':time.monotonic()-tick,'environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_DEBUG','CARGO_PROFILE_DEV_DEBUG','CARGO_PROFILE_TEST_DEBUG','TMPDIR','RUSTDOCFLAGS']},'exit_code':result.returncode,'source_unchanged':unchanged,'source_manifest_sha256':sha(HERE/'source-manifest.json'),'log_sha256':sha(HERE/f'{name}.log')})
    print(name,result.returncode,flush=True)
    assert result.returncode == 0 and unchanged
