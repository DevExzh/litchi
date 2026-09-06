#!/usr/bin/env python3
"""Freeze the ODT three-role protocol after passing candidate pilots."""
import datetime
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def write(path, value):
    with path.open('x') as output:
        output.write(json.dumps(value, indent=2) + '\n')

def main():
    for name in ('after-buffered-pilots', 'after-streaming-pilots'):
        receipt = json.loads((ROOT / 'checks' / (name + '.json')).read_text())
        assert receipt['status'] == 'pass' and receipt['source_unchanged'] is True
    (ROOT / 'oracle').mkdir(exist_ok=False)
    for source, target in [('protocol-draft.json', 'protocol.json'), ('verify-report.py', 'verify-report.py')]:
        (ROOT / 'oracle' / target).write_bytes((ROOT / source).read_bytes())
    protocol = json.loads((ROOT / 'protocol-draft.json').read_text())
    protocol['status'] = 'frozen'
    protocol['classification'] = 'Matched ODT three-role fresh paragraph creation; descriptive latency and memory tradeoff evidence.'
    protocol['oracle'] = {
        'path': 'oracle/protocol.json', 'sha256': sha(ROOT / 'oracle/protocol.json'),
        'verifier_path': 'oracle/verify-report.py', 'verifier_sha256': sha(ROOT / 'oracle/verify-report.py'),
    }
    protocol['profile_scope'] = {
        'roles': ['before-buffered', 'after-buffered', 'after-streaming'],
        'mode': 'normal', 'shape': 'large', 'captures': ['perf-stat', 'perf-record'],
        'events': ['cycles:u', 'instructions:u', 'branches:u', 'branch-misses:u', 'L1-dcache-load-misses:u'],
        'call_graph': 'fp', 'required': True,
    }
    write(ROOT / 'protocol.json', protocol)
    spec = importlib.util.spec_from_file_location('verify0435freeze', ROOT / 'verify.py')
    verifier = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(verifier)
    verifier.validate_protocol(protocol)
    write(ROOT / 'frozen-inputs.json', {
        'status': 'frozen', 'utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
        'files': {name: sha(ROOT / name) for name in (
            'protocol.json', 'oracle/verify-report.py', 'oracle/protocol.json', 'check.py')},
    })
    print('Frozen 36 reports, 1080 samples, six profiles; three roles, three shapes, two modes and repeats')

if __name__ == '__main__':
    main()
