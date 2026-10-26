#!/usr/bin/env python3
"""Discard only generated debug caches, between commands, when space is tight."""
from pathlib import Path
import shutil
import subprocess
import sys
path = Path(sys.argv[1]).resolve()
if path.name != 'target' or not path.is_dir():
    raise SystemExit(f'expected existing target directory, got {path}')
size = int(subprocess.check_output(['du', '-s', '-B1', str(path)], text=True).split()[0])
if size > 16 * 1024**3:
    debug = path / 'debug'
    print(f'pruning generated debug cache ({size} bytes in target): {debug}', flush=True)
    shutil.rmtree(debug, ignore_errors=False)
else:
    print(f'target_bytes={size}', flush=True)
