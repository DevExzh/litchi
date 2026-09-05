#!/usr/bin/env python3
"""Replay all four debug reports through the final portable report checker."""
import ast
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent


def main():
    for path in sorted(ROOT.glob('*.py')):
        ast.parse(path.read_text(), filename=str(path))
    reports = sorted((ROOT / 'smoke').glob('*.json'))
    reports = [path for path in reports if path.name != 'commands.json']
    assert len(reports) == 4
    for path in reports:
        result = subprocess.check_output([sys.executable, '-B', str(ROOT / 'probe-report.py'), str(path)])
        row = json.loads(result)
        assert row['status'] == 'pass' and len(row['mutations']) == 20
        print(json.dumps({'report': path.name, **row}))


if __name__ == '__main__':
    main()
