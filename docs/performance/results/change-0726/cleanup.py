#!/usr/bin/env python3
"""Remove only the three owned 0726 external roots, preserving executable identities."""
import hashlib,json,shutil
from pathlib import Path
P=Path(__file__).resolve().parent
roots=[Path('/home/zhuhe/code')/name for name in ('litchi-target-0726','litchi-0726-bin','litchi-0726-profile')]
assert not (P/'cleanup.json').exists(), 'cleanup already recorded'
identities=[]
for root in roots:
 assert root.is_dir() and not root.is_symlink(), root
 if 'target-' not in root.name:
  for path in sorted(root.rglob('*')):
   if path.is_file():
    h=hashlib.sha256()
    with path.open('rb') as f:
     for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
    identities.append(dict(path=str(path),bytes=path.stat().st_size,sha256=h.hexdigest()))
receipt=dict(roots=[str(p) for p in roots],identities=identities,removed=False)
(P/'cleanup.json').write_text(json.dumps(receipt,indent=2)+'\n')
for root in roots:shutil.rmtree(root)
for path in P.rglob('__pycache__'):shutil.rmtree(path)
receipt['removed']=all(not p.exists() for p in roots)
assert receipt['removed']
(P/'cleanup.json').write_text(json.dumps(receipt,indent=2)+'\n')
print('PASS removed three owned roots; retained',len(identities),'file identity witnesses')
