#!/usr/bin/env python3
"""Remove only this packet's build/binary scratch after exact identity checks."""
import hashlib,json,shutil
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
assert read(P/'analysis.json')['status']=='passed'
assert read(P/'audit.json')['status']=='passed' and read(P/'negative-checks.json')['status']=='passed'
assert read(P/'disposition.json')['status'] in ['retain','reject']
roots=[ROOT.parent/'litchi-target-0735',ROOT.parent/'litchi-0735-bin'];identities=[]
for v in ['baseline','candidate']:
 for b in read(P/f'{v}-build.json')['binaries'].values():
  f=Path(b['path']);assert f.parent==roots[1] and not f.is_symlink();assert f.stat().st_size==b['bytes'] and sha(f)==b['sha256'];identities.append(b)
assert {p.name for p in roots[1].iterdir()}=={Path(b['path']).name for b in identities}
for root in roots:
 assert root.parent==ROOT.parent and not root.is_symlink();shutil.rmtree(root)
for cache in P.rglob('__pycache__'):shutil.rmtree(cache)
(P/'cleanup.json').write_text(json.dumps(dict(removed=True,binaries=identities,roots=[str(r) for r in roots]),indent=2)+'\n')
print('PASS removed only owned 0735 target and four exact-identity binaries')
