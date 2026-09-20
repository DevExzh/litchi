#!/usr/bin/env python3
"""Seal or verify the exact terminal evidence packet after final gates."""
import hashlib,json,sys
from pathlib import Path
P=Path(__file__).resolve().parent
manifest=P/'artifact-manifest.json'
current={str(p.relative_to(P)):{'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
         for p in sorted(P.rglob('*')) if p.is_file() and p!=manifest}
assert not any('__pycache__' in p for p in current)
if sys.argv[1:]==['--write']:
    manifest.write_text(json.dumps({'files':current},indent=2)+'\n')
elif sys.argv[1:]==['--check']:
    assert json.loads(manifest.read_text())['files']==current,'artifact census or hashes changed'
else:raise SystemExit('usage: artifact-seal.py --write|--check')
print('PASS:',len(current),'exact packet artifacts')
