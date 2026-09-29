"""Remove only the two exclusively created, marker-bound 0834 temporary roots."""
import shutil,time
import driver_v3 as d
import audit

def main():
 d.check('repaired-v3')
 assert d.read(d.P/'reader-replay.json')['status']=='pass'
 assert d.read(d.P/'reader-replay.json')['formal_samples']==0
 for name in ['reader-tests.json','pptx-reader-tests.json','opc-reader-tests.json']:
  assert d.read(d.P/name)['status']=='pass'
 assert audit.build_audit()['status']=='pass'
 binaries=[d.read(d.P/f'build-{stage}.json')['binary'] for stage in ['diagnostic','repaired-v2','repaired-v3']]
 for binary in binaries:assert d.desc(binary['path'])==binary
 roots=[]
 for root in [d.TARGET,d.SCRATCH]:
  assert root.is_dir() and not root.is_symlink()
  assert d.read(root/'.owner-0834.json')==dict(packet=str(d.P),path=str(root))
  files=[p for p in root.rglob('*') if p.is_file() or p.is_symlink()]
  roots.append(dict(path=str(root),files=len(files),logical_bytes=sum(p.lstat().st_size for p in files),marker=d.desc(root/'.owner-0834.json')))
 start=time.time();d.write(d.P/'cleanup-started.json',dict(started_unix=start,binaries=binaries,roots=roots))
 for root in [d.TARGET,d.SCRATCH]:shutil.rmtree(root)
 assert not d.TARGET.exists() and not d.SCRATCH.exists()
 d.write(d.P/'cleanup.json',dict(status='pass',started_unix=start,finished_unix=time.time(),binaries=binaries,roots=roots,target_absent=True,scratch_absent=True))
 print('cleanup PASS:',sum(r['files'] for r in roots),'files,',sum(r['logical_bytes'] for r in roots),'logical bytes')
if __name__=='__main__':main()
