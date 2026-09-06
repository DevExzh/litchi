#!/usr/bin/env python3
"""Capture the frozen normal baseline pilots before production source changes."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def main():
    frozen = json.loads((ROOT / 'frozen-inputs.json').read_text())
    for name, digest in frozen['files'].items():
        assert sha(ROOT / name) == digest
    binary = json.loads((ROOT / 'before/binary-copies.json').read_text())['normal']
    assert sha(Path(binary['path'])) == binary['sha256']
    directory = ROOT / 'before/pilots'
    directory.mkdir()
    for shape in ('tiny', 'medium', 'large'):
        report = directory / (shape + '.json')
        catalog = directory / (shape + '-catalog.json')
        resource = directory / (shape + '-resource.log')
        log = directory / (shape + '.log')
        argv = ['taskset', '-c', '2', '/usr/bin/time', '-v', '-o', str(resource),
            binary['path'], '--case', 'ods_streaming_create', '--semantic-shape', shape,
            '--workers', '1', '--samples', '30', '--warmup', '3', '--json', str(report),
            '--corpus-manifest', str(catalog)]
        with log.open('xb') as stream:
            subprocess.run(argv, stdout=stream, stderr=subprocess.STDOUT, check=True)
        subprocess.run([sys.executable, '-B', str(ROOT / 'oracle/verify-report.py'),
            '--report', str(report), '--mode', 'normal', '--shape', shape,
            '--role', 'after-streaming'], check=True)
        receipt = {'status': 'pass', 'shape': shape, 'argv': argv, 'binary': binary,
            'driver_sha256': sha(Path(__file__)), 'frozen_inputs': frozen,
            'build_receipt': 'checks/before-build.json',
            'artifacts': {str(p.relative_to(ROOT)): {'bytes': p.stat().st_size,
                'sha256': sha(p)} for p in (report, catalog, resource, log)}}
        (directory / (shape + '-receipt.json')).write_text(json.dumps(receipt, indent=2)+'\n')
        print('PASS baseline pilot', shape, flush=True)

if __name__ == '__main__':
    main()
