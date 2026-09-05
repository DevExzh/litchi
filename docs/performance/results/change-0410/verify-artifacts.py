#!/usr/bin/env python3
"""Verify the published 0410 inventory and lossless compressed payloads."""
import hashlib
import json
from pathlib import Path
import subprocess

root = Path(__file__).resolve().parent
manifest = json.loads((root / 'artifact-manifest.json').read_text())
for item in manifest['files']:
    path = root / item['path']
    payload = path.read_bytes()
    if len(payload) != item['bytes'] or hashlib.sha256(payload).hexdigest() != item['sha256']:
        raise SystemExit(f'Artifact mismatch: {path}')
for item in manifest['compressed_originals']:
    path = root / item['compressed_path']
    subprocess.run(['zstd', '-q', '-t', str(path)], check=True)
    payload = subprocess.check_output(['zstd', '-q', '-dc', str(path)])
    if len(payload) != item['bytes'] or hashlib.sha256(payload).hexdigest() != item['sha256']:
        raise SystemExit(f'Compressed original mismatch: {path}')
print(f"Verified {len(manifest['files'])} published files and "
      f"{len(manifest['compressed_originals'])} compressed originals")
