#!/usr/bin/env python3
"""Inventory retained evidence; exclude reproducible source and disposable targets."""
import argparse, hashlib, json, os
from pathlib import Path
ap=argparse.ArgumentParser();ap.add_argument('--root',type=Path,required=True);ap.add_argument('--output',type=Path,required=True);a=ap.parse_args()
excluded={'source-baseline','source-candidate','target-quality','target-candidate-managed-v1','__pycache__'}
rows=[]
for dp,dns,fns in os.walk(a.root):
 dns[:]=sorted(x for x in dns if x not in excluded)
 for n in sorted(fns):
  p=Path(dp)/n
  if p.resolve()==a.output.resolve():continue
  if p.is_symlink():rows.append({'path':str(p.relative_to(a.root)),'symlink':os.readlink(p)});continue
  h=hashlib.sha256()
  with p.open('rb') as f:
   for b in iter(lambda:f.read(1<<20),b''):h.update(b)
  rows.append({'path':str(p.relative_to(a.root)),'bytes':p.stat().st_size,'mode':p.stat().st_mode&0o777,'sha256':h.hexdigest()})
result={'root':str(a.root),'excluded_directories':sorted(excluded),'source_directory_contents':'See retained source-manifest-baseline-v1.json and source-manifest-candidate-v1.json, inventoried below','files':rows,'file_count':len(rows),'total_bytes':sum(x.get('bytes',0) for x in rows)}
with a.output.open('x') as f:json.dump(result,f,indent=2);f.write('\n')
print(json.dumps({'files':len(rows),'bytes':result['total_bytes'],'output':str(a.output)}))
