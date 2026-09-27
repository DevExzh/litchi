"""Remove only the owned 0788 target after verifying all four executable identities."""
import os,shutil
from pathlib import Path
import custody as c
origin=c.read(c.P/'origin.json')
assert str(c.TARGET)==origin['target']
assert c.TARGET.name=='litchi-target-0788' and c.TARGET.is_dir() and not c.TARGET.is_symlink()
assert c.source()['files']==c.read(c.P/'restored-source.json')['files']
assert not (c.P/'cleanup.json').exists()
binaries=[]
for leg in ('before','after'):
 for artifact in c.read(c.P/f'build-{leg}/build.json')['binaries'].values():
  assert Path(artifact['path']).parent==c.TARGET
  assert c.artifact(artifact['path'])==artifact
  binaries.append(artifact)
assert len(binaries)==4 and len({a['path'] for a in binaries})==4
size=sum(p.stat().st_size for p in c.TARGET.rglob('*') if p.is_file())
shutil.rmtree(c.TARGET)
assert not c.TARGET.exists()
for cache in c.P.rglob('__pycache__'):
 assert cache.is_dir() and not cache.is_symlink();shutil.rmtree(cache)
c.write(c.P/'cleanup.json',{'schema':'litchi.cached-part-memory-cleanup.0788.v1','verified':True,'executables_verified_before_removal':True,'binaries':binaries,'target':str(c.TARGET),'target_bytes_before_removal':size,'target_absent_after_removal':True,'source_restored_files':9196,'worktree_removal':'Owned worktree, branch, copied root lock and three exact third-party symlinks are removed after evidence commit integration; main report records closure.'})
print('verified four binaries; removed',size,'owned target bytes')
