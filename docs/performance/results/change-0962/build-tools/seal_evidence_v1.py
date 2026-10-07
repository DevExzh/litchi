#!/usr/bin/env python3
"""Commit bounded source/metadata copies, inventory retained external artifacts."""
from pathlib import Path
import hashlib,json,shutil,datetime
ROOT=Path('/home/zhuhe/litchi-performance-0962')
DEST=Path('/home/zhuhe/code/litchi/docs/performance/results/change-0962')
def digest(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''):h.update(b)
 return h.hexdigest()
def main():
 assert not DEST.exists()
 rows=[];copies=[]
 for p in sorted(ROOT.rglob('*')):
  if not p.is_file() or '__pycache__' in p.parts:continue
  assert not p.is_symlink(),p
  rel=p.relative_to(ROOT);size=p.stat().st_size
  row={'path':str(rel),'bytes':size,'sha256':digest(p)}
  copy=size<1048576 and p.suffix in {'.json','.py','.txt','.log','.md'} and 'executables' not in rel.parts
  row['bounded_copy']=copy;rows.append(row)
  if copy:copies.append((p,DEST/rel))
 DEST.mkdir(parents=True)
 for src,dst in copies:
  dst.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(src,dst)
 manifest={'schema':'litchi.performance.0962.external-artifact-inventory.v1','created_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'external_root':str(ROOT),'files':rows,'regeneration':'Exact build/capture commands, environment and source identities are in retained receipts/start records and source helpers. Commands target create-only paths: use a fresh external root or restore the declared layout. Decode original perf.data with the recorded paired perf-script commands. Failed DWARF profiles remain diagnostic only; invalid build attempts are not regeneration authority.'}
 encoded=(json.dumps(manifest,indent=2)+'\n').encode();assert len(encoded)<1048576
 (DEST/'external-artifact-inventory.json').write_bytes(encoded)
 print(json.dumps({'files':len(rows),'bounded_copies':len(copies),'external_bytes':sum(r['bytes'] for r in rows),'max_committed_file':max(p.stat().st_size for p in DEST.rglob('*') if p.is_file())}))
if __name__=='__main__':main()
