#!/usr/bin/env python3
"""Build one role inside its clean worktree and retain its binary identity."""
import datetime
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from capture import ROOT, TEMP, FLAGS, sha


def main():
    role = sys.argv[1]
    assert role in ('control', 'candidate')
    tree = TEMP / (role + '-tree')
    revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=tree, text=True).strip()
    assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=tree).strip()
    target = ROOT.parents[3] / 'tools/perf-baseline/target'
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', RUSTFLAGS=FLAGS, CARGO_PROFILE_RELEASE_DEBUG='1',
               CARGO_INCREMENTAL='0', CARGO_BUILD_JOBS='4', DEBUGINFOD_URLS='', LC_ALL='C')
    argv = ['cargo', 'build', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
            '--target-dir', str(target), '--bin', 'litchi-perf-baseline']
    record = dict(role=role, revision=revision, argv=argv, cwd=str(tree), clean_before=True,
                  environment={k:env[k] for k in ['RUSTUP_TOOLCHAIN','RUSTFLAGS','CARGO_PROFILE_RELEASE_DEBUG','CARGO_INCREMENTAL','CARGO_BUILD_JOBS','DEBUGINFOD_URLS']},
                  started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    with (ROOT / (role+'-build.stdout')).open('x') as out, (ROOT / (role+'-build.stderr')).open('x') as err:
        result = subprocess.run(argv, cwd=tree, env=env, stdout=out, stderr=err)
    assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=tree).strip()
    record.update(exit_code=result.returncode, clean_after=True, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    (ROOT / (role+'-build.json')).write_text(json.dumps(record, indent=2)+'\n')
    if result.returncode:
        raise SystemExit(result.returncode)
    binary = TEMP / role
    shutil.copy2(target / 'release/litchi-perf-baseline', binary)
    binding = dict(revision=revision, binary_sha256=sha(binary), bytes=binary.stat().st_size,
                   build_receipt=role+'-build.json', build_receipt_sha256=sha(ROOT/(role+'-build.json')),
                   clean_build=True)
    (ROOT / (role+'-binding.json')).write_text(json.dumps(binding, indent=2)+'\n')
    print(json.dumps(binding), flush=True)


if __name__ == '__main__':
    main()
