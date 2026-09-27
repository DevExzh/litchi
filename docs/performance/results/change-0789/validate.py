"""Offline custody, qualification and final-seal validator; never runs probe."""
import gzip,json,re,subprocess,sys
import common as c


def verify_binary(a):
 if c.resolve(a).exists():return c.verify(a)
 cleanup=c.read(c.P/'cleanup.json')
 assert cleanup['target_removed'] and cleanup['target']==str(c.TARGET)
 assert a in cleanup['removed_binaries'] and not c.TARGET.exists()
 return c.resolve(a)


def validate():
 original=(c.P/'capture.py').read_text().split('\ndef main():')[0]
 detailed=(c.P/'capture-detail.py').read_text().split('\ndef main():')[0]
 assert detailed==original.replace("['/usr/bin/strace','-f','-qq'", "['/usr/bin/strace','-v','-f','-qq'")
 frozen=c.read(c.P/'frozen-inputs.json')
 for name,sha in frozen.items():assert c.sha(c.P/name)==sha,name
 production=c.read(c.P/'production-source.json')['files']
 for name,sha in production.items():assert c.sha(c.ROOT/name)==sha,name
 for name,sha in c.read(c.P/'architecture-inputs.json').items():assert c.sha(c.ROOT/name)==sha,name
 if '--check-workspace' in sys.argv:
  origin=c.read(c.P/'origin.json')
  for name,sha in origin['unrelated'].items():assert c.sha(c.ROOT/name)==sha,name
  worktrees=subprocess.check_output(['git','worktree','list','--porcelain'],cwd=c.ROOT,text=True)
  # Main HEAD advances when committed; every other worktree remains exact.
  assert worktrees.split('\n\n',1)[1]==origin['worktrees'].split('\n\n',1)[1]
 build=c.read(c.P/'build.json');assert len(build['commands'])==2
 for command in build['commands']:assert command['exit_code']==0;assert not c.verify(command['log']).read_bytes()
 for kind in ['binary','sanitized_binary']:verify_binary(build[kind])
 c.verify(build['inputs'])
 checks=c.read(c.P/'qualification/checks.json');assert len(checks)==9
 for r in checks:
  verify_binary(r['binary']);out=c.verify(r['stdout']).read_text();err=c.verify(r['stderr']).read_text()
  if r['expected_success']:
   assert r['exit_code']==0 and not err and r['stdin_hex']==(b'+\n'*6).hex()
   lines=out.splitlines();assert len(lines)==12
   mib=int(r['command'][2]);pages=mib*1024*1024//4096;checksum=sum((i%251)+1 for i in range(pages))
   for i,phase in enumerate(c.read(c.P/'plan.json')['phases']):
    before=lines[2*i].split('\t');after=lines[2*i+1].split('\t')
    assert len(before)==8 and len(after)==6 and before[:2]==['RSS0789',phase] and after[:3]==['ACK0789',phase,before[2]]
    assert int(before[6])==(mib*1024*1024 if i in [1,2,3] else 0)
    assert int(before[7])==(0 if i<2 else checksum)
  else:assert r['exit_code']>0 and err
 for lane,count in [('matrix',272),('traces',4),('trace-detail',4)]:
  done=c.read(c.P/lane/'complete.json');assert done['children']==count
  receipts=c.read(c.verify(done['receipts']));assert len(receipts)==count
  for a in receipts:
   r=c.read(c.verify(a));assert r['error'] is None and r['wait4']['exit_code']==0
   verify_binary(r['binary'])
   points=json.loads(gzip.decompress(c.verify(r['points']).read_bytes()))
   allowed=c.read(c.P/'plan.json')['affinities'][r['affinity']]
   assert r['command'][:3]==['/usr/bin/taskset','-c',','.join(map(str,allowed))]
   assert r['command'][-7:]==[build['binary']['path'],'--mib',str(r['case']['mib']),'--workers',str(r['case']['workers']),'--touch',r['case']['touch']]
   for point in points:
    stat=point['raw']['stat'];tail=stat[stat.rfind(')')+2:].split()
    assert int(tail[36]) in allowed
    assert point['identity']['exe']==build['binary']['path']
    if r['observer']=='full':
     mask=re.search(r'^Cpus_allowed_list:\s+(.+)$',point['raw']['status'],re.M)[1].strip()
     assert mask==('0' if r['affinity']=='single' else '0-31')
   for key in ['points','transcript','stderr','gnu_time','strace']:
    if key in r:c.verify(r[key])
 for name in ['analyze.py','accounting_summary.py']:
  subprocess.run([sys.executable,'-B',str(c.P/name),'--check'],check=True)
 if '--require-final-seal' in sys.argv:
  seal=c.read(c.P/'seal.json');actual={str(p.relative_to(c.P)) for p in c.P.rglob('*') if p.is_file() and p.name!='seal.json'}
  assert actual==set(seal['files'])
  for name,a in seal['files'].items():assert a['path']==name;c.verify(a)
  for a in seal['documents']:c.verify(a)
  assert not c.TARGET.exists() and not list(c.P.rglob('__pycache__'))
 print(f'0789 validation PASS: {len(production)} production files, 35 architecture inputs, nine qualifications, 280 children')

if __name__=='__main__':validate()
