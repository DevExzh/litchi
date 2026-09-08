#!/usr/bin/env python3
"""Seal every regular bundle file, excluding only the seal itself."""
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent
rows = []
for path in sorted(ROOT.rglob('*')):
    assert not path.is_symlink(), path
    assert '__pycache__' not in path.parts and path.suffix != '.pyc', path
    if not path.is_file() or path == ROOT / 'SHA256SUMS':
        continue
    with path.open('rb') as stream:
        digest = hashlib.file_digest(stream, 'sha256').hexdigest()
    rows.append(f'{digest}  {path.relative_to(ROOT).as_posix()}\n')
(ROOT / 'SHA256SUMS').write_text(''.join(rows))
print(f'sealed {len(rows)} files')
