#!/usr/bin/env python3
"""Reproduce the four corrupt-report checks without modifying the capture."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import tempfile

p = argparse.ArgumentParser()
p.add_argument('--root', type=Path, required=True)
p.add_argument('--repo-root', type=Path, required=True)
p.add_argument('--output', type=Path, required=True)
a = p.parse_args()
root = a.root.resolve()
verifier = Path(__file__).resolve().parents[1] / 'verify-capture.py'
results = []
for name, filename in [('duplicate-order', 'normal-1.json'),
                       ('duplicate-metric-index', 'allocator-1.json'),
                       ('opaque-read', 'normal-2.json'),
                       ('wrong-oracle', 'normal-3.json')]:
    source = root / filename
    raw = (source.read_bytes() if source.exists() else
           subprocess.check_output(['zstd', '-q', '-dc', str(source) + '.zst']))
    report = json.loads(raw)
    if name == 'duplicate-order':
        order = report['results'][0]['elapsed_ns']['sample_order']
        order[0] = order[1]
    elif name == 'duplicate-metric-index':
        indices = report['results'][0]['operation_metrics']['sample_indices']
        indices[0] = indices[1]
    elif name == 'opaque-read':
        report['results'][1]['source']['xls']['opaque_payload_read_bytes'][0] = 1
    else:
        report['results'][0]['output_sha256'] = '0' * 64
    with tempfile.TemporaryDirectory(prefix='litchi-goal-0411-negative-') as tmp:
        directory = Path(tmp)
        for artifact in root.iterdir():
            if artifact.is_file():
                if artifact.suffix == '.zst':
                    # zstd refuses symlink inputs by default. Keep compressed
                    # fixtures as private regular files for the replay.
                    shutil.copyfile(artifact, directory / artifact.name)
                else:
                    (directory / artifact.name).symlink_to(artifact)
        # Never write through a symlink into the original capture.
        target = directory / filename
        target.unlink(missing_ok=True)
        target.write_text(json.dumps(report))
        run = subprocess.run([
            'python3', str(verifier), '--root', str(directory),
            '--repo-root', str(a.repo_root.resolve()),
            '--output', str(directory / 'negative-result.json'),
        ], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        marker = {
            'duplicate-order': 'sample_order must be a complete permutation',
            'duplicate-metric-index': 'sample_indices',
            'opaque-read': 'opaque_payload_read_bytes',
            'wrong-oracle': 'output_sha256',
        }[name]
        if run.returncode == 0 or marker not in run.stdout:
            raise SystemExit(f'Expected rejection not observed: {name}\n{run.stdout}')
        results.append(dict(mutation=name, original_sha256=hashlib.sha256(raw).hexdigest(),
                            rejected=True, exit_code=run.returncode, diagnostic=run.stdout))
a.output.write_text(json.dumps(results, indent=2) + '\n')
print('Four corrupted reports rejected')
