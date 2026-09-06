#!/usr/bin/env python3
import hashlib,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
paths=sorted(p for p in ROOT.rglob('*') if p.is_file() and p.name!='SHA256SUMS')
assert not list(ROOT.rglob('__pycache__'))
assert all(not p.is_symlink() for p in ROOT.rglob('*'))
(ROOT/'SHA256SUMS').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+str(p.relative_to(ROOT))+'\n' for p in paths))
print(json.dumps({'status':'pass','members':len(paths)}))
