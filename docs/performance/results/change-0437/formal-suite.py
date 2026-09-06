#!/usr/bin/env python3
"""Run the frozen ODP matrix and profiles serially through source custody."""
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def run(tag, driver, *args):
    command = [sys.executable, '-B', str(ROOT / 'check.py'), '--tag', tag, '--',
               sys.executable, '-B', str(ROOT / driver), *args,
               '--repo-root', str(REPO), '--protocol', str(ROOT / 'protocol.json'),
               '--custody-driver', str(ROOT / 'check.py')]
    print(json.dumps({'tag': tag, 'argv': command}), flush=True)
    subprocess.run(command, cwd=REPO, check=True)


def main():
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    phases = protocol['matrix']['phases']
    assert phases == ['A1', 'B1', 'C1', 'C2', 'B2', 'A2']
    for phase in phases:
        run('formal-' + phase.lower(), 'capture.py', '--phase', phase, '--attempt', 'formal')
    for role in ('before-buffered', 'after-buffered', 'after-streaming'):
        for kind in ('stat', 'record'):
            run(f'formal-{role}-{kind}', 'profile.py', '--role', role, '--kind', kind, '--attempt', 'formal')


if __name__ == '__main__':
    main()
