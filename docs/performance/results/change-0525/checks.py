"""Run XLSX reconstruction correctness and policy gates in one serial lane."""
import argparse
from run import TARGET, run


COMMANDS = [
    ('fmt', ['cargo', 'fmt', '--all', '--check']),
    ('harness-fmt', ['cargo', 'fmt', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--all', '--check']),
    ('xlsx-tests', ['cargo', 'test', '--locked', '-p', 'litchi-xlsx', '--all-features', '--', '--test-threads=2']),
    ('harness-noncompact-test', ['cargo', 'test', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', 'xlsx_noncompact_cell_reference_guard']),
    ('harness-test', ['cargo', 'test', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', 'xlsx_source_cell_values_commit_allocation_metrics_align_with_phases']),
    ('harness-existing-metrics-tests', ['cargo', 'test', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', 'xlsx_commit_metrics_tests']),
    ('workspace-check', ['cargo', 'check', '--locked', '--workspace', '--all-features']),
    ('clippy', ['cargo', 'clippy', '--locked', '-p', 'litchi-xlsx', '--all-features', '--lib', '--', '-D', 'warnings']),
    ('harness-clippy', ['cargo', 'clippy', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', '--', '-D', 'warnings']),
    ('rustdoc', ['cargo', 'doc', '--locked', '-p', 'litchi-xlsx', '--all-features', '--no-deps']),
    ('boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py']),
    ('claims', ['python3', '-B', 'tools/check_perf_claims.py', '--registry', 'docs/performance/claim-registry-v1.json', '--repo-root', '.', '--evidence-root', '.', '--mode', 'strict']),
]


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--stage', choices=['candidate', 'final'], default='candidate')
    args = parser.parse_args()
    for name, command in COMMANDS:
        settings = ['env', 'CARGO_TARGET_DIR='+str(TARGET), 'CARGO_BUILD_JOBS=2',
                    'CARGO_INCREMENTAL=0', 'RUSTDOCFLAGS=-D warnings']
        run(args.stage, 'check-'+name, settings+command)
