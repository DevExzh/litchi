"""Remove only marker-bound 0836 roots after retained evidence passes."""
from pathlib import Path
import gzip,hashlib,shutil,time
import driver as d

def main():
 for stage in ['baseline','fp']:d.check(stage)
 for name in ['profile-admission.json','reader-tests.json','profile-tests.json','profile-analysis.json']:
  assert d.read(d.P/name)['status']=='pass',name
 assert d.read(d.P/'validation/independent-final-preview/receipt.json')['exit_code']==0
 descriptors={}
 def walk(value):
  if isinstance(value,dict):
   if all(k in value for k in ['path','bytes','sha256']):
    path=Path(value['path'])
    if path.is_relative_to(d.TARGET) or path.is_relative_to(d.SCRATCH):
     descriptor={k:value[k] for k in ['path','bytes','sha256']}
     assert d.desc(path)==descriptor,path
     if str(path) in descriptors:assert descriptors[str(path)]==descriptor
     descriptors[str(path)]=descriptor
   for child in value.values():walk(child)
  elif isinstance(value,list):
   for child in value:walk(child)
 for path in d.P.rglob('*.json'):walk(d.read(path))
 compressed=[]
 for manifest in ['profiles.json','traces.json','mappings.json']:
  for row in d.read(d.P/manifest)['rows']:
   for raw,packed in [('raw','raw_gzip'),('decoded_plain','decoded')]:
    if packed not in row:continue
    h=hashlib.sha256();size=0
    with gzip.open(row[packed]['path'],'rb') as f:
     for chunk in iter(lambda:f.read(1048576),b''):h.update(chunk);size+=len(chunk)
    assert h.hexdigest()==row[raw]['sha256'] and size==row[raw]['bytes']
    compressed.append(dict(original=row[raw],retained=row[packed]))
 roots=[]
 for root in [d.TARGET,d.SCRATCH]:
  assert root.is_dir() and not root.is_symlink()
  assert d.read(root/'.owner-0836.json')==dict(packet=str(d.P),path=str(root))
  files=[p for p in root.rglob('*') if p.is_file() or p.is_symlink()]
  roots.append(dict(path=str(root),files=len(files),logical_bytes=sum(p.lstat().st_size for p in files),marker=d.desc(root/'.owner-0836.json')))
 start=time.time()
 witness=dict(started_unix=start,roots=roots,owned_paths=[str(d.TARGET),str(d.SCRATCH)],artifacts=list(descriptors.values()),compressed=compressed)
 d.write(d.P/'cleanup-started.json',witness)
 for root in [d.TARGET,d.SCRATCH]:shutil.rmtree(root)
 assert not d.TARGET.exists() and not d.SCRATCH.exists()
 d.write(d.P/'cleanup.json',dict(**witness,status='pass',finished_unix=time.time(),target_absent=True,scratch_absent=True))
 print('cleanup PASS:',sum(r['files'] for r in roots),'files,',sum(r['logical_bytes'] for r in roots),'logical bytes')

if __name__=='__main__':main()
