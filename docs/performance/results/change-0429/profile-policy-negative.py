#!/usr/bin/env python3
"""Reject the supplementary sample count outside its two declared roles."""
import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def main():
    profile = json.loads((ROOT / 'profiles/media-rich-bytes.json').read_text())
    native = json.loads((ROOT / 'native-cli-control.json').read_text())
    cases = []
    plain = copy.deepcopy(profile)
    plain['corpus'] = 'plain'
    cases.append(('plain', plain))
    ranged = copy.deepcopy(profile)
    ranged['provider'] = 'range'
    cases.append(('range', ranged))
    native.update(samples=100, warmup=3, checked_iteration_count=103)
    cases.append(('native', native))
    with tempfile.TemporaryDirectory(prefix='litchi-0429-profile-policy-') as directory:
        for name, row in cases:
            path = Path(directory) / (name + '.json')
            path.write_text(json.dumps(row))
            result = subprocess.run([sys.executable, '-B', str(ROOT / 'verify-report.py'), str(path)], capture_output=True, text=True)
            assert result.returncode != 0 and '100-sample profile reports require' in result.stderr, (name, result.stderr)
    print(json.dumps({'status': 'pass', 'rejected_profile_roles': [name for name, _ in cases]}))


if __name__ == '__main__':
    main()
