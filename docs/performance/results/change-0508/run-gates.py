#!/usr/bin/env python3
"""Run final harness gates sequentially after the descriptive captures."""
from pathlib import Path
import subprocess
import sys

here = Path(__file__).resolve().parent
for lane in ['tests', 'clippy', 'rustdoc', 'doctests', 'check-features']:
    result = subprocess.run([sys.executable, '-B', str(here / 'run.py'), lane])
    if result.returncode:
        raise SystemExit(result.returncode)
