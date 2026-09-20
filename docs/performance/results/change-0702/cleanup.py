#!/usr/bin/env python3
"""Remove only this batch's named scratch after frozen evidence has passed."""
import hashlib,json,shutil,sys
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
assert sys.argv[1:]==['--apply']
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
lock=ROOT/'Cargo.lock'
lock_hash=sha(lock) if lock.exists() else None
targets=[ROOT.parent/'litchi-target-0702',ROOT.parent/'litchi-0702-bin',ROOT.parent/'litchi-0702-profile',P/'marker-control.pptx']
for phase in ['baseline','candidate']:
 for row in json.loads((P/f'builds-{phase}.json').read_text())+[json.loads((P/f'build-refusal-{phase}.json').read_text()),json.loads((P/f'build-oracle-{phase}.json').read_text())]:
  assert sha(Path(row['binary']))==row['binary_sha256']
records=[]
for target in targets:
 assert not target.is_symlink()
 record=dict(path=str(target),existed=target.exists(),kind='directory' if target.is_dir() else 'file')
 if target.is_dir():shutil.rmtree(target)
 elif target.exists():
  record['sha256']=sha(target);target.unlink()
 record['removed']=not target.exists();records.append(record)
assert (sha(lock) if lock.exists() else None)==lock_hash
(P/'cleanup.json').write_text(json.dumps(dict(removed=records,workspace_cargo_lock_preserved_sha256=lock_hash,scope='Exact batch-owned scratch only; shared targets and unrelated work preserved.'),indent=2)+'\n')
print('removed exactly four owned scratch paths; workspace Cargo.lock preserved')
