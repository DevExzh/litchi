"""Root-only serial quality gates and fresh binaries; never overwrite attempts."""
import os
from pathlib import Path
import subprocess
import time
import custody as c


if __name__ == '__main__':
    p = c.P
    target = Path(c.read(p / 'plan.json')['target'])
    source = c.census()
    i = 0
    while (p / f'quality-{i}').exists():
        i += 1
    out = p / f'quality-{i}'
    out.mkdir()
    c.write(out / 'source.json', source)
    c.write(p / 'source.json', source)
    m = ['--manifest-path', 'tools/perf-baseline/Cargo.toml']
    locked = ['--offline', '--locked']
    allocation = ['--features', 'allocator-metrics']
    commands = [
        ['cargo', 'fmt', *m, '--', '--check'],
        ['cargo', 'check', *m, *locked, *allocation, '--all-targets'],
        ['cargo', 'test', *m, *locked, *allocation, '--bins', '--', '--test-threads=1'],
        ['cargo', 'clippy', *m, *locked, *allocation, '--lib', '--bin', 'ordinary_save_artifacts', '--', '-D', 'warnings'],
        ['cargo', 'doc', *m, *locked, *allocation, '--no-deps'],
        ['python3', '-B', 'tools/test_perf_baseline_source_policy.py'],
        ['python3', '-B', 'tools/validate_crud_coverage_index.py'],
        ['cargo', 'build', *m, *locked, '--release', '--bin', 'litchi-perf-baseline', '--bin', 'ordinary_save_artifacts'],
        ['cargo', 'build', *m, *locked, '--release', *allocation, '--bin', 'litchi-perf-baseline-alloc'],
    ]
    env = os.environ | {'CARGO_TARGET_DIR': str(target), 'CARGO_BUILD_JOBS': '2', 'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_DEV_DEBUG': '0', 'RUSTDOCFLAGS': '-D warnings', 'PYTHONDONTWRITEBYTECODE': '1'}
    rows = []
    binaries = {}
    for index, command in enumerate(commands):
        log = out / f'{index:02}.log'
        started = time.time()
        with log.open('w') as f:
            run = subprocess.run(command, cwd=c.ROOT, env=env, stdout=f, stderr=subprocess.STDOUT)
        rows.append({'command': command, 'exit_code': run.returncode, 'started': started, 'ended': time.time(), 'log': str(log.relative_to(p)), 'log_sha256': c.sha(log)})
        c.write(out / 'checks.json', rows)
        print(f'gate {index+1}/{len(commands)} exit {run.returncode}', flush=True)
        assert run.returncode == 0, rows[-1]
        assert c.census() == source, 'source changed during quality/build'
        if index == 7:
            binaries['native'] = c.artifact(target / 'release/litchi-perf-baseline')
            binaries['export'] = c.artifact(target / 'release/ordinary_save_artifacts')
        if index == 8:
            binaries['allocation'] = c.artifact(target / 'release/litchi-perf-baseline-alloc')
    for binary in binaries.values():
        assert c.artifact(Path(binary['path'])) == binary
    c.write(p / 'build.json', {'source_sha256': c.sha(p / 'source.json'), 'binaries': binaries, 'rows': rows, 'target': str(target), 'environment': {k: env.get(k) for k in ['CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL', 'CARGO_PROFILE_DEV_DEBUG', 'RUSTDOCFLAGS', 'RUSTFLAGS', 'RUSTUP_TOOLCHAIN']}})
