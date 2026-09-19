#!/usr/bin/env python3
"""Build expanded refusal coverage at exact baseline, then restore frozen candidate.

This runs only under the coordinator's exclusive source/build lane. It archives
both source states and restores candidate bytes in finally, like trace.py.
"""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
base=json.loads((P/'baseline.json').read_text())
allowed={f'crates/litchi-pptx/src/{name}' for name in ['parts/mod.rs','parts/slide.rs','presentation/package.rs','opened/model.rs','opened/tests.rs','notes/codec.rs','notes/package.rs','notes/mod.rs']}
def sha(b):return hashlib.sha256(b).hexdigest()
changes=[]
for name,digest in base['source_sha256'].items():
 current=(ROOT/name).read_bytes()
 if sha(current)!=digest:
  assert name in allowed,name
  old=subprocess.check_output(['git','show',f"{base['baseline_head']}:{name}"],cwd=ROOT)
  assert sha(old)==digest
  changes.append((name,current,old))
assert changes
out=P/'expanded-refusal-baseline';out.mkdir(exist_ok=False)
for name,current,old in changes:
 stem=name.replace('/','__')
 (out/(stem+'.candidate')).write_bytes(current)
 (out/(stem+'.baseline')).write_bytes(old)
manifest=dict(status='started',baseline_head=base['baseline_head'],sources=[dict(path=n,candidate_sha256=sha(c),baseline_sha256=sha(b)) for n,c,b in changes])
def save():(out/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
save()
try:
 for name,current,old in changes:
  assert (ROOT/name).read_bytes()==current
  (ROOT/name).write_bytes(old)
 command=['python3',str(P/'build-refusal.py'),'baseline']
 result=subprocess.run(command,cwd=ROOT)
 manifest.update(command=command,exit_code=result.returncode)
 assert result.returncode==0
 manifest['status']='built'
finally:
 restored=[]
 for name,current,old in reversed(changes):
  assert (ROOT/name).read_bytes()==old, f'unexpected concurrent source mutation: {name}'
  (ROOT/name).write_bytes(current)
  restored.append(dict(path=name,restored_exact=(ROOT/name).read_bytes()==current))
 manifest['restoration']=restored
 manifest['restored_exact']=all(r['restored_exact'] for r in restored)
 save()
print('expanded refusal baseline built; candidate sources restored exactly')
