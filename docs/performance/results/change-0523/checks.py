"""Serial quality gates for the measurement enabler and unchanged CFB owner."""
from run import TARGET, run

COMMANDS = [
    ('fmt', ['cargo', 'fmt', '--all', '--check']),
    ('harness-fmt', ['cargo', 'fmt', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--all', '--check']),
    ('cfb-tests', ['cargo', 'test', '--locked', '-p', 'litchi-cfb', '--all-features']),
    ('cfb-no-default-tests', ['cargo', 'test', '--locked', '-p', 'litchi-cfb', '--no-default-features']),
    ('workspace-check', ['cargo', 'check', '--locked', '--workspace', '--all-features']),
    ('cfb-clippy', ['cargo', 'clippy', '--locked', '-p', 'litchi-cfb', '--all-features', '--lib', '--', '-D', 'warnings']),
    ('harness-clippy', ['cargo', 'clippy', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', '--', '-D', 'warnings']),
    ('harness-rustdoc', ['cargo', 'doc', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', '--no-deps']),
    ('boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py']),
    ('claims', ['python3', '-B', 'tools/check_perf_claims.py', '--registry', 'docs/performance/claim-registry-v1.json', '--repo-root', '.', '--evidence-root', '.', '--mode', 'strict']),
]

if __name__ == '__main__':
    for name, command in COMMANDS:
        run('check-' + name, ['env', 'CARGO_TARGET_DIR=' + str(TARGET),
            'CARGO_BUILD_JOBS=2', 'CARGO_INCREMENTAL=0', 'RUSTDOCFLAGS=-D warnings'] + command)
