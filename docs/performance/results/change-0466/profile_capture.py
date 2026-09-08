#!/usr/bin/env python3
"""Same-source diagnostic build with explicit frame pointers and unwind tables."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import capture

BUNDLE = Path(__file__).resolve().parent
REPO = capture.REPO


def run(argv, path, env):
    receipt = dict(argv=argv, environment={k: env.get(k) for k in
                   ('RUSTUP_TOOLCHAIN', 'RUSTFLAGS', 'CARGO_PROFILE_RELEASE_DEBUG', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'DEBUGINFOD_URLS')},
                   started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    with path.with_suffix('.stdout').open('x') as stdout, path.with_suffix('.stderr').open('x') as stderr:
        result = subprocess.run(argv, cwd=REPO, env=env, stdout=stdout, stderr=stderr)
    receipt.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    path.with_suffix('.json').write_text(json.dumps(receipt, indent=2) + '\n')
    print(path.name, result.returncode, flush=True)
    if result.returncode:
        raise SystemExit(result.returncode)


def main():
    binding = capture.source_check()
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4', CARGO_INCREMENTAL='0',
               CARGO_PROFILE_RELEASE_DEBUG='1', RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
               DEBUGINFOD_URLS='', LC_ALL='C', PYTHONDONTWRITEBYTECODE='1')
    run(['cargo', 'build', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
         '--bin', 'litchi-perf-baseline'], BUNDLE / 'profile-build', env)
    capture.source_check()
    binary = capture.TEMP / 'profile'
    shutil.copy2(REPO / 'tools/perf-baseline/target/release/litchi-perf-baseline', binary)
    identity = dict(binary_sha256=capture.sha(binary), bytes=binary.stat().st_size,
                    source_manifest_sha256=binding['source']['sha256'],
                    revision=subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO, text=True).strip(),
                    scope='same source as normal, different profiling build flags; not normal latency evidence')
    (BUNDLE / 'profile-binding.json').write_text(json.dumps(identity, indent=2) + '\n')
    output = BUNDLE / 'samples-fp'
    output.mkdir()
    workload = [str(binary), '--case', 'xlsx_one_percent_commit_save', '--xlsx-shape', 'dense-wide',
                '--workers', '1', '--warmup', '3', '--samples', '50', '--json', str(output / 'report.json'),
                '--corpus-manifest', str(output / 'corpus-catalog.json')]
    run(['taskset', '-c', '2', '/usr/bin/time', '-v', '-o', str(output / 'resource.log'),
         'perf', 'record', '-F', '499', '-e', 'cycles:u', '--call-graph', 'fp',
         '-o', str(output / 'perf.data'), '--'] + workload, output / 'capture', env)
    capture.source_check()
    assert capture.sha(binary) == identity['binary_sha256']
    run(['perf', 'script', '--no-inline', '-i', str(output / 'perf.data'),
         '-F', 'comm,pid,tid,time,event,period,ip,sym,dso'], output / 'perf-script', env)
    run(['perf', 'report', '--stdio', '--no-children', '--no-inline', '-g', 'none',
         '-i', str(output / 'perf.data'), '--sort', 'symbol', '--percent-limit', '0.5'], output / 'top-symbols', env)


if __name__ == '__main__':
    main()
