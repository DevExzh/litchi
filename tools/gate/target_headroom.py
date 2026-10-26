#!/usr/bin/env python3
"""Discard only generated debug caches, between commands, when space is tight.

The target may hold at most `GATE_TARGET_MAX_GB` GiB (default 200) before its
`debug` cache is discarded. Files that vanish while the size is measured (for
example a concurrent cargo run removing them) are skipped, not fatal.
"""
import os
from pathlib import Path
import shutil
import sys


def target_bytes(path: Path) -> int:
    total = 0
    for root, _dirs, files in os.walk(path, onerror=lambda _error: None):
        for name in files:
            try:
                total += os.lstat(os.path.join(root, name)).st_blocks * 512
            except FileNotFoundError:
                continue
    return total


def main() -> int:
    path = Path(sys.argv[1]).resolve()
    if path.name != 'target' or not path.is_dir():
        raise SystemExit(f'expected existing target directory, got {path}')
    limit = int(os.environ.get('GATE_TARGET_MAX_GB', '200')) * 1024**3
    size = target_bytes(path)
    if size > limit:
        debug = path / 'debug'
        print(f'pruning generated debug cache ({size} bytes in target): {debug}', flush=True)
        shutil.rmtree(debug, ignore_errors=True)
    else:
        print(f'target_bytes={size}', flush=True)
    return 0


if __name__ == '__main__':
    sys.exit(main())
