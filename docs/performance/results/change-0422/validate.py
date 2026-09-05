#!/usr/bin/env python3
"""Run serialized focused validation and retain exact commands and raw output."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
output = ROOT / 'validation.json'
if output.exists():
    raise SystemExit('refusing to overwrite existing validation receipt')
manifest = ['--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml']
test = ['cargo', '+1.98.1', 'test', *manifest, '--features', 'allocator-metrics']
clippy = ['cargo', '+1.98.1', 'clippy', *manifest, '--all-features', '--all-targets', '--no-deps', '--']
checks = [
    ('counter-tests', [*test, '--lib', 'allocation_metrics::tests', '--', '--test-threads=1'], {}, True),
    ('aggregation-tests', [*test, '--lib', 'operation_metrics::tests', '--', '--test-threads=1'], {}, True),
    ('filesystem-serialization', [*test, '--lib', 'filesystem::tests::child_output_keeps_allocation_sample', '--', '--test-threads=1'], {}, True),
    ('allocator-wrapper', [*test, '--bin', 'litchi-perf-baseline-alloc', '--', '--test-threads=1'], {}, True),
    ('comparator-tests', [sys.executable, '-m', 'unittest', 'discover', '-s', 'tools', '-p', 'test_perf_compare.py'], {}, True),
    ('clippy-strict', [*clippy, '-D', 'warnings'], {}, False),
    ('clippy-diagnostic', [*clippy, '-A', 'clippy::chunks_exact_to_as_chunks', '-A', 'clippy::clone_on_copy', '-A', 'clippy::needless_lifetimes'], {}, True),
    ('rustdoc', ['cargo', '+1.98.1', 'doc', *manifest, '--features', 'allocator-metrics', '--lib', '--no-deps'], {'RUSTDOCFLAGS': '-D warnings'}, True),
]
paths = ['tools/perf-baseline/src/allocation_metrics.rs', 'tools/perf-baseline/src/operation_metrics.rs', 'tools/perf-baseline/src/filesystem.rs', 'tools/perf_compare.py', 'tools/test_perf_compare.py']
def source_hashes():
    return {p: hashlib.sha256((REPO / p).read_bytes()).hexdigest() for p in paths}
record = {'change': 422, 'status': 'running', 'source_sha256_before': source_hashes(), 'checks': [], 'scope': 'focused counter, aggregation, serialization, real allocator wrapper and comparator tests; strict lint debt reported separately'}
env = os.environ | {'RUSTUP_TOOLCHAIN': '1.98.1', 'PYTHONDONTWRITEBYTECODE': '1'}
for name, argv, extra, required in checks:
    item = {'name': name, 'argv': argv, 'environment_overrides': extra, 'required_for_this_batch': required, 'started_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'log': name + '.log'}
    with (ROOT / item['log']).open('wb') as log:
        result = subprocess.run(argv, cwd=REPO, env=env | extra, stdout=log, stderr=subprocess.STDOUT)
    item.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    record['checks'].append(item)
    output.write_text(json.dumps(record, indent=2, sort_keys=True) + '\n')
    print(name, result.returncode, flush=True)
    if required and result.returncode:
        record['status'] = 'failed'
        break
else:
    record['status'] = 'pass_with_strict_lint_debt' if any(x['exit_code'] for x in record['checks']) else 'pass'
record['source_sha256_after'] = source_hashes()
if record['source_sha256_after'] != record['source_sha256_before']:
    record['status'] = 'failed_source_changed'
output.write_text(json.dumps(record, indent=2, sort_keys=True) + '\n')
raise SystemExit(0 if record['status'] in ['pass', 'pass_with_strict_lint_debt'] else 1)
