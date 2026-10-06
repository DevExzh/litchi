from __future__ import annotations
import hashlib
import json
import os
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path('/tmp/litchi-performance-0925')
SOURCE_ROOT = ROOT / 'measurement-source-v4'
BINARY = ROOT / 'measurement-target-v4-final/release/litchi-perf-baseline'
RECEIPT = ROOT / 'measurement-build-v4-final.json'
OUT = ROOT / 'release-preflight-v4'
OUT.mkdir(exist_ok=True)
ROUTES = ['generic_read_at', 'owned_vec', 'owned_reader', 'file_warm', 'range_short', 'range_zero', 'range_delayed']
CASE = 'docx_source_backed_existing_section_layout_edit_save'

def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()

def now() -> str:
    return datetime.now(timezone.utc).isoformat()

binary_sha = sha256(BINARY)
source_file = SOURCE_ROOT / 'tools/perf-baseline/src/lib.rs'
source_sha = sha256(source_file)
receipt_sha = sha256(RECEIPT)
results = []
for route in ROUTES:
    report = OUT / f'{route}.json'
    stdout = OUT / f'{route}.stdout.log'
    stderr = OUT / f'{route}.stderr.log'
    argv = [str(BINARY), '--case', CASE, '--warmup', '1', '--samples', '3', '--json', str(report)]
    env = os.environ.copy()
    env['LITCHI_DOCX_INGRESS_ROUTE'] = route
    started = now()
    monotonic = time.monotonic()
    with stdout.open('wb') as out, stderr.open('wb') as err:
        try:
            proc = subprocess.run(argv, cwd=str(SOURCE_ROOT), env=env, stdout=out, stderr=err, timeout=900)
            exit_code = proc.returncode
            timed_out = False
        except subprocess.TimeoutExpired:
            exit_code = None
            timed_out = True
    ended = now()
    item = {
        'route': route,
        'argv': argv,
        'cwd': str(SOURCE_ROOT),
        'environment': {'LITCHI_DOCX_INGRESS_ROUTE': route},
        'started_utc': started,
        'ended_utc': ended,
        'elapsed_wall_seconds': time.monotonic() - monotonic,
        'exit_code': exit_code,
        'timed_out': timed_out,
        'report': {'path': str(report), 'bytes': report.stat().st_size, 'sha256': sha256(report)} if report.is_file() else None,
        'stdout': {'path': str(stdout), 'bytes': stdout.stat().st_size, 'sha256': sha256(stdout)},
        'stderr': {'path': str(stderr), 'bytes': stderr.stat().st_size, 'sha256': sha256(stderr)},
    }
    results.append(item)
    print(json.dumps(item, sort_keys=True), flush=True)
    if exit_code != 0 or timed_out or item['report'] is None:
        break

receipt = {
    'schema': 'litchi.performance.0925.provider-release-preflight.v1',
    'status': 'completed' if len(results) == len(ROUTES) and all(x['exit_code'] == 0 and x['report'] for x in results) else 'failed',
    'source_root': str(SOURCE_ROOT),
    'source_file': str(source_file),
    'source_sha256': source_sha,
    'binary': {'path': str(BINARY), 'bytes': BINARY.stat().st_size, 'sha256': binary_sha},
    'build_receipt': {'path': str(RECEIPT), 'bytes': RECEIPT.stat().st_size, 'sha256': receipt_sha},
    'case': CASE,
    'warmup': 1,
    'samples': 3,
    'routes': ROUTES,
    'results': results,
    'created_utc': now(),
}
path = OUT / 'receipt.json'
path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print(json.dumps({'receipt': str(path), 'status': receipt['status']}, sort_keys=True), flush=True)
raise SystemExit(0 if receipt['status'] == 'completed' else 1)
