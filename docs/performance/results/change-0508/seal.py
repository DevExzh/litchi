#!/usr/bin/env python3
"""Seal retained evidence; temporary build products are intentionally excluded."""
import hashlib
from pathlib import Path

HERE = Path(__file__).resolve().parent
rows = []
for path in sorted(HERE.rglob('*')):
    if not path.is_file() or '__pycache__' in path.parts or path.name == 'SHA256SUMS':
        continue
    assert not path.is_symlink()
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    rows.append(f'{digest}  {path.relative_to(HERE)}\n')
(HERE / 'SHA256SUMS').write_text(''.join(rows))
print(f'sealed {len(rows)} evidence files')
