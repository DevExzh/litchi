#!/usr/bin/env python3
"""Seal retained artifacts without relying on the checkout or executables."""
from pathlib import Path
from common import sha


def seal(root):
    files = sorted(p for p in root.rglob('*') if p.is_file() and p.name != 'SHA256SUMS')
    assert all(not p.is_symlink() and '__pycache__' not in p.parts for p in files)
    (root / 'SHA256SUMS').write_text(''.join(
        f'{sha(p)}  {p.relative_to(root).as_posix()}\n' for p in files))


if __name__ == '__main__':
    seal(Path(__file__).resolve().parent)
