#!/usr/bin/env python3
"""Seal the complete current bundle, excluding only the root seal itself."""
import hashlib
from pathlib import Path

root = Path(__file__).resolve().parent
seal = root / 'SHA256SUMS'
rows = []
for path in sorted(root.rglob('*')):
    if path.is_symlink():
        raise RuntimeError(f'symlink in evidence bundle: {path}')
    if path.is_file() and path != seal:
        if '__pycache__' in path.parts:
            raise RuntimeError(f'bytecode cache in evidence bundle: {path}')
        rows.append(f'{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}\n')
seal.write_text(''.join(rows))
print(f'Sealed {len(rows)} files.')
