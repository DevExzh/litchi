#!/usr/bin/env python3
"""Verify every published 0413 artifact and each lossless original."""
import hashlib
import json
from pathlib import Path
import subprocess

root = Path(__file__).resolve().parent
manifest = json.loads((root / 'artifact-manifest.json').read_text())
expected = {item['path'] for item in manifest['files']}
actual = {str(p.relative_to(root)) for p in root.rglob('*') if p.is_file()}
if actual != expected | {'artifact-manifest.json'}:
    raise SystemExit('Published file inventory differs from the manifest')
for item in manifest['files']:
    payload = (root / item['path']).read_bytes()
    if len(payload) != item['bytes'] or hashlib.sha256(payload).hexdigest() != item['sha256']:
        raise SystemExit(f"Artifact mismatch: {item['path']}")
for item in manifest['compressed_originals']:
    path = root / item['compressed_path']
    payload = subprocess.check_output(['zstd', '-q', '-dc', str(path)])
    if len(payload) != item['bytes'] or hashlib.sha256(payload).hexdigest() != item['sha256']:
        raise SystemExit(f"Compressed original mismatch: {path}")
print(f"Verified {len(expected)} published artifacts and "
      f"{len(manifest['compressed_originals'])} compressed originals")
