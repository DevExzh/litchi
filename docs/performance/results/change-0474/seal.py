#!/usr/bin/env python3
"""Seal exact regular-file coverage of this standalone evidence bundle."""
import hashlib
from pathlib import Path

root = Path(__file__).resolve().parent
rows = []
for path in sorted(root.rglob('*')):
    assert not path.is_symlink(), path
    assert '__pycache__' not in path.parts, path
    if path.is_file() and path != root / 'SHA256SUMS':
        with path.open('rb') as stream:
            digest = hashlib.file_digest(stream, 'sha256').hexdigest()
        rows.append(f'{digest}  {path.relative_to(root).as_posix()}\n')
(root / 'SHA256SUMS').write_text(''.join(rows))
print('sealed', len(rows), 'regular files')
