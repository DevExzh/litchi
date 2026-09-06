#!/usr/bin/env python3
"""Bind the ODT same-API ABBA protocol after passing both sets of pilots."""
import datetime
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    for role in ('before', 'after'):
        receipt = json.loads((ROOT / 'checks' / (role + '-pilots.json')).read_text())
        assert receipt['status'] == 'pass' and receipt['source_unchanged'] is True
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert protocol['status'] == 'frozen'
    for path, digest in [('oracle/protocol.json', protocol['oracle']['sha256']), ('oracle/verify-report.py', protocol['oracle']['verifier_sha256'])]:
        assert sha(ROOT / path) == digest
    spec = importlib.util.spec_from_file_location('verify0436freeze', ROOT / 'verify.py')
    verifier = importlib.util.module_from_spec(spec); spec.loader.exec_module(verifier)
    verifier.validate_protocol(protocol)
    names = ['protocol.json', 'oracle/protocol.json', 'oracle/verify-report.py', 'check.py']
    with (ROOT / 'frozen-inputs.json').open('x') as output:
        output.write(json.dumps({'status': 'frozen', 'utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'files': {name: sha(ROOT / name) for name in names}}, indent=2) + '\n')
    print('Frozen 24 reports, 720 samples, four profiles; same streaming API before and after')

if __name__ == '__main__':
    main()
