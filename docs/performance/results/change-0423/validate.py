#!/usr/bin/env python3
"""Serialize focused 0423 checks and retain source-bound commands and output."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--tag', default='validation')
parser.add_argument('--group', choices=['rust', 'repository', 'regression', 'followup', 'lint'], default='rust')
args = parser.parse_args()
output = ROOT / (args.tag + '.json')
if output.exists():
    raise SystemExit('refusing to overwrite validation receipt')
manifest = ['--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml']
test = ['cargo', '+1.98.1', 'test', *manifest, '--features', 'allocator-metrics', '--lib']
clippy = ['cargo', '+1.98.1', 'clippy', *manifest, '--all-features', '--all-targets', '--no-deps', '--']
checks = [
    ('cross-copy-tests', [*test, 'cross_copy', '--', '--test-threads=1'], {}, True),
    ('clippy-strict', [*clippy, '-D', 'warnings'], {}, False),
    ('clippy-diagnostic', [*clippy, '-A', 'clippy::chunks_exact_to_as_chunks', '-A', 'clippy::clone_on_copy', '-A', 'clippy::needless_lifetimes'], {}, True),
    ('rustdoc', ['cargo', '+1.98.1', 'doc', *manifest, '--features', 'allocator-metrics', '--lib', '--no-deps'], {'RUSTDOCFLAGS': '-D warnings'}, True),
] if args.group in ('rust', 'followup', 'lint') else [
    ('claims', [sys.executable, 'tools/check_perf_claims.py', '--registry', 'docs/performance/claim-registry-v1.json', '--repo-root', '.', '--evidence-root', '.', '--mode', 'strict'], {}, True),
    ('boundaries', [sys.executable, 'tools/check_crate_boundaries.py'], {}, True),
    ('crud-index', [sys.executable, 'tools/validate_crud_coverage_index.py'], {}, True),
]
if args.group == 'lint':
    checks = checks[1:]
if args.group == 'followup':
    checks[0:1] = [
        ('source-lifecycle', [*test, 'tests::pptx_source_backed_cross_copy_lifecycles_cover_plain_and_media_oracles', '--', '--exact', '--test-threads=1'], {}, True),
        ('media-oracle-regression', [*test, 'tests::pptx_source_backed_cross_copy_media_lifecycle_gate_rejects_tampered_media_relationship_geometry_and_count', '--', '--exact', '--test-threads=1'], {}, True),
    ]
if args.group == 'regression':
    checks = [('media-oracle-regression', [*test, 'pptx_source_backed_cross_copy_media_lifecycle_gate', '--', '--test-threads=1'], {}, True)]
paths = ['tools/perf-baseline/src/lib.rs', 'tools/validate_crud_coverage_index.py', 'tools/test_crud_coverage_index.py', 'docs/performance/crud-coverage-index-v1.json']
def source_hashes():
    return {p: hashlib.sha256((REPO / p).read_bytes()).hexdigest() for p in paths}
def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()
record = {'change': 423, 'status': 'running', 'source_sha256_before': source_hashes(), 'checks': [], 'scope': args.group}
env = os.environ | {'RUSTUP_TOOLCHAIN': '1.98.1', 'PYTHONDONTWRITEBYTECODE': '1'}
for name, argv, extra, required in checks:
    log_path = ROOT / (args.tag + '-' + name + '.log')
    if log_path.exists():
        raise SystemExit('refusing existing log: ' + str(log_path))
    item = {'name': name, 'argv': argv, 'environment_overrides': extra, 'required_for_this_batch': required, 'started_utc': now(), 'log': log_path.name}
    with log_path.open('wb') as log:
        result = subprocess.run(argv, cwd=REPO, env=env | extra, stdout=log, stderr=subprocess.STDOUT)
    item.update(exit_code=result.returncode, finished_utc=now())
    if name in ('source-lifecycle', 'media-oracle-regression'):
        item['nonzero_test_selection'] = 'running 1 test' in log_path.read_text()
        if not item['nonzero_test_selection']:
            item['exit_code'] = 1
    record['checks'].append(item)
    output.write_text(json.dumps(record, indent=2, sort_keys=True) + '\n')
    print(name, item['exit_code'], flush=True)
    if required and item['exit_code']:
        record['status'] = 'failed'
        break
else:
    record['status'] = 'pass_with_strict_lint_debt' if any(x['exit_code'] for x in record['checks']) else 'pass'
record['source_sha256_after'] = source_hashes()
if record['source_sha256_after'] != record['source_sha256_before']:
    record['status'] = 'failed_source_changed'
output.write_text(json.dumps(record, indent=2, sort_keys=True) + '\n')
raise SystemExit(0 if record['status'] in ['pass', 'pass_with_strict_lint_debt'] else 1)
