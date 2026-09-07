#!/usr/bin/env python3
"""Reject a resealed one-nanosecond mutation of the derived timing summary."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    receipt = ROOT / 'tamper-test.json'
    if receipt.exists():
        raise RuntimeError('tamper-test.json already exists')
    with tempfile.TemporaryDirectory(prefix='litchi-0464-tamper-') as directory:
        copied = Path(directory) / 'bundle'
        shutil.copytree(ROOT, copied)
        summary = copied / 'summary.json'
        before = sha(summary)
        value = json.loads(summary.read_text())
        metric = value['lanes'][0]['timings']['api_sum_ns']
        original = metric['p50']
        metric['p50'] += 1
        summary.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')
        rows = []
        for path in sorted(copied.rglob('*')):
            if path.is_symlink():
                raise RuntimeError('symlink in copied bundle')
            if path.is_file() and path.name != 'SHA256SUMS':
                rows.append(f'{sha(path)}  {path.relative_to(copied).as_posix()}\n')
        (copied / 'SHA256SUMS').write_text(''.join(rows))
        argv = [sys.executable, '-B', str(copied / 'verify.py')]
        result = subprocess.run(argv, capture_output=True, text=True)
        combined = result.stdout + result.stderr
        expected = 'summary.json is not the deterministic recomputation'
        if result.returncode == 0 or expected not in combined:
            raise RuntimeError(f'mutation did not fail for recomputation: {combined}')
        record = {
            'schema': 'litchi-0464-summary-tamper-v1', 'status': 'pass',
            'driver_sha256': sha(Path(__file__)),
            'verifier_sha256': sha(ROOT / 'verify.py'),
            'original_summary_sha256': before,
            'mutation': {'path': 'lanes[0].timings.api_sum_ns.p50',
                         'before': original, 'after': metric['p50']},
            'resealed': True, 'argv': argv, 'exit_code': result.returncode,
            'stdout': result.stdout, 'stderr': result.stderr,
        }
    record['temporary_directory_absent'] = not Path(directory).exists()
    with receipt.open('x') as output:
        json.dump(record, output, indent=2)
        output.write('\n')
    print(json.dumps({'status': 'pass', 'resealed_mutation_rejected': True}))


if __name__ == '__main__':
    main()
