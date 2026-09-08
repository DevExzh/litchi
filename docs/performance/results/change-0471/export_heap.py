#!/usr/bin/env python3
"""Export whole-process Heaptrack summaries; invoke under the capture CPU lock."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

for lane in ('A-heap', 'B-heap'):
    output = ROOT / lane
    paths = sorted(p for p in output.glob('heaptrack*') if p.suffix in ('.gz', '.zst'))
    assert len(paths) == 1, paths
    argv = ['heaptrack_print', '-f', str(paths[0]), '-n', '0', '-a', '0', '-p', '0', '-T', '0', '-l', '0']
    receipt = dict(schema='litchi-0471-heap-export-v1', argv=argv, cwd=str(ROOT),
                   input_path=paths[0].name, input_sha256=sha(paths[0]),
                   driver_sha256=sha(Path(__file__)),
                   started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    with (output / 'heaptrack-print.stdout').open('xb') as stdout, (output / 'heaptrack-print.stderr').open('xb') as stderr:
        result = subprocess.run(argv, cwd=ROOT, env=dict(os.environ, LC_ALL='C', DEBUGINFOD_URLS=''), stdout=stdout, stderr=stderr)
    receipt.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                   artifacts={p.name: dict(sha256=sha(p), bytes=p.stat().st_size) for p in output.glob('heaptrack-print.std*')})
    with (output / 'heaptrack-print.json').open('x') as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True)
        stream.write('\n')
    print(lane, result.returncode, flush=True)
    if result.returncode:
        raise SystemExit(result.returncode)
