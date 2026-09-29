"""Remove only the two exclusively owned 0837 roots after successful closure."""
import os,shutil,time
from pathlib import Path
import driver as d
d.check()
assert d.read(d.P/'closure.json')['status']=='pass'
assert d.read(d.P/'final-v2-quality.json')['status']=='pass'
assert d.read(d.P/'red-green.json')['status']=='pass'
artifacts=[d.read(d.P/'build-baseline.json')['binary']]
for value in artifacts:assert d.desc(value['path'])==value
roots=[]
for root in [d.TARGET,d.SCRATCH]:
 assert root.is_dir() and not root.is_symlink()
 assert root.resolve()==root and d.read(root/'.owner-0837.json')==dict(packet=str(d.P),path=str(root))
 files=0;size=0
 for directory,_,names in os.walk(root):
  for name in names:
   files+=1;size+=(Path(directory)/name).lstat().st_size
 roots.append(dict(path=str(root),files=files,logical_bytes=size))
d.write(d.P/'cleanup-started.json',dict(started_unix=time.time(),roots=roots,artifacts=artifacts))
for root in [d.TARGET,d.SCRATCH]:shutil.rmtree(root)
assert not d.TARGET.exists() and not d.SCRATCH.exists()
d.write(d.P/'cleanup.json',dict(status='pass',finished_unix=time.time(),roots=roots,artifacts=artifacts))
print('cleanup PASS',sum(r['files'] for r in roots),'files',sum(r['logical_bytes'] for r in roots),'logical bytes')
