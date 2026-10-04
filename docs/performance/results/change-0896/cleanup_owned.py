#!/usr/bin/env python3
"""Remove only the marked 0896 isolated source and build trees after review."""
from pathlib import Path
import datetime, hashlib, json, os, shutil
P=Path(__file__).resolve().parent
R=P.parents[3]
S=R/'target/litchi-0896-source'
T=R/'target/litchi-0896-build'
def sha(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
 return h.hexdigest()
def desc(p):return {'path':str(p),'bytes':p.stat().st_size,'sha256':sha(p)}
def main():
 output=P/'owned-cleanup.json'
 assert not output.exists()
 for name in ['verification.json','native-review.json']:
  assert json.loads((P/name).read_text())['status']=='pass',name
 decision=json.loads((P/'decision.json').read_text())
 assert decision['status'] in ['rejected','adopted']
 for p in [S,T]:
  assert not p.is_symlink() and p.is_dir()
  assert (p/'.owner-litchi-0896').read_text()=='0896 owned source/build\n'
 for proc in Path('/proc').iterdir():
  if not proc.name.isdigit() or int(proc.name)==os.getpid():continue
  try:
   exe=(proc/'exe').resolve();cwd=(proc/'cwd').resolve()
  except (OSError,RuntimeError):continue
  assert not any(q==base or base in q.parents for q in [exe,cwd] for base in [S,T]),str(proc)
 binaries=json.loads((P/'candidate-build.json').read_text())['binaries']
 for d in binaries.values():
  p=Path(d['path']);assert p.stat().st_size==d['bytes'] and sha(p)==d['sha256']
 census=json.loads((P/'source-after-build.json').read_text())['files']
 direct=json.loads((P/'direct-plan.json').read_text())['fixtures']
 fixture_names={d['relative'] for d in direct.values()}|{'test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx'}
 relocated={}
 for name in sorted(fixture_names):
  before=desc(S/name);after=desc(R/name)
  assert before['sha256']==after['sha256']==census[name]['sha256']
  assert before['bytes']==after['bytes']==census[name]['bytes']
  relocated[name]={'original':before,'retained':after}
 removed=[]
 for p in [S,T]:
  size=sum(f.stat().st_size for f in p.rglob('*') if f.is_file() and not f.is_symlink())
  shutil.rmtree(p)
  removed.append({'path':str(p),'bytes':size,'absent':not p.exists()})
 result={'status':'pass','utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'removed':removed,'fixture_relocation':relocated,'retained_release_binaries':binaries,'evidence':{n:desc(P/n) for n in ['decision.json','source-after-build.json','native-review.json','verification.json','reports-cleanup.json']},'source_reconstruction':'Retained 0894 source archive + 0896 candidate overlay, support receipts and quality-Cargo.lock. Direct fixture paths under removed source map to identical repository test-data paths and source manifest hashes.','worktrees_created':[],'unowned_worktrees':'preserved','helper':desc(Path(__file__))}
 with output.open('x') as f:json.dump(result,f,indent=2,sort_keys=True);f.write('\n')
 print('Owned cleanup PASS',sum(r['bytes'] for r in removed))
if __name__=='__main__':main()
