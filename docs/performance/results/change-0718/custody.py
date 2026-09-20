#!/usr/bin/env python3
"""Shared source census and paths for native and procfs diagnostic builds."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0718'
BIN = ROOT.parent / 'litchi-0718-bin'

def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def census():
    files = [ROOT/'Cargo.toml', ROOT/'Cargo.lock', *(ROOT/'.cargo').rglob('*')]
    files += [p for folder in ['crates','tools/perf-baseline'] for p in (ROOT/folder).rglob('*')
              if 'target' not in p.parts and (p.suffix == '.rs' or p.name in ['Cargo.toml','Cargo.lock'])]
    return {str(p.relative_to(ROOT)):sha(p) for p in sorted(files) if p.is_file()}
