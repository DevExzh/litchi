#!/usr/bin/env python3
"""Replay all actual preflight controls with the final adversarial report probes."""
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent


def main():
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    rows = []
    for lane in protocol['lanes']:
        name = lane['scenario'] + '-' + lane['corpus']
        path = ROOT / 'smoke-verified' / (name + '.json')
        result = json.loads(subprocess.check_output([
            sys.executable, '-B', str(ROOT / 'probe-report.py'), str(path)]))
        assert result['status'] == 'pass'
        rows.append({'name': name, **result})
    print(json.dumps({'status': 'pass', 'reports': rows,
                      'mutations': sum(len(row['mutations']) for row in rows)}))


if __name__ == '__main__':
    main()
