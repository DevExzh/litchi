"""Root custody audit for the OPC caller-thread census packet."""
import argparse,subprocess,sys
from pathlib import Path
import custody as c

def artifact(a):
 f=Path(a['path']);assert f.is_file() and not f.is_symlink() and c.artifact(f)==a,f
 return f

def validate(seal=False):
 p=c.P;base=c.read(p/'source.json');origin=c.read(p/'origin.json');plan=c.read(p/'plan.json')
 assert base['revision']==origin['base'] and c.source()['files']==base['files']
 other=lambda s:[b for b in s.strip().split('\n\n') if not b.startswith('worktree '+str(c.ROOT)+'\n')]
 assert other(subprocess.check_output(['git','worktree','list','--porcelain'],cwd=c.ROOT,text=True))==other(origin['worktrees'])
 for n,h in origin['unrelated'].items():assert c.sha(c.ROOT/n)==h
 for n,h in c.read(p/'architecture-inputs.json').items():assert c.sha(c.ROOT/n)==h
 assert c.sha(p/'workspace-Cargo.lock')==c.sha(c.ROOT/'Cargo.lock')
 for a in c.read(p/'inheritance.json').values():artifact(a)
 q=c.read(p/'quality/complete.json');last=0;assert len(q['rows'])==3
 for r in q['rows']:
  assert r['exit_code']==0 and last<=r['started']<=r['ended'];last=r['ended'];artifact(r['log'])
 for n,h in q['inputs'].items():assert c.sha(p/'hook-test-src'/n)==h
 assert c.sha(p/'hook-test-src/src/xml_attributes.rs')==c.sha(p/'instrumentation/after/xml_attributes.rs')
 binaries=[]
 for leg in ['before','after']:
  b=c.read(p/('build-'+leg)/'build.json');source=c.read(artifact(b['source']));assert source['revision']==base['revision']
  changed={n:h for n,h in source['files'].items() if h!=base['files'][n]}
  assert set(changed)==(set() if leg=='before' else set(plan['source_allowlist']))
  if leg=='after':assert changed['crates/litchi-opc/src/xml_attributes.rs']==c.sha(p/'instrumentation/after/xml_attributes.rs')
  inputs=c.read(artifact(b['inputs']))
  for n,h in inputs['frozen'].items():assert c.sha(p/n)==h
  for n,h in inputs['probe'].items():assert c.sha(p/'probe-src'/n)==h
  for n,h in inputs['instrumentation'].items():assert c.sha(p/'instrumentation'/n)==h
  assert len(b['rows'])==4 and b['environment']=={'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','RUSTFLAGS':None}
  for r in b['rows']:
   assert r['exit_code']==0 and last<=r['started']<=r['ended'];last=r['ended'];artifact(r['log'])
  binaries.append(b['binary'])
 if c.TARGET.exists():
  for b in binaries:artifact(b)
 else:
  cleanup=c.read(p/'cleanup.json');assert cleanup['target_removed'] and cleanup['target']==str(c.TARGET) and cleanup['removed_binaries']==binaries
 restoration=c.read(p/'restored-source.json');assert restoration['files']==base['files']
 reports=0
 for lane,count,index in [('control',15,0),('census',30,1)]:
  done=c.read(p/lane/'complete.json');rows=c.read(artifact(done['receipts']));assert done['children']==len(rows)==count
  assert c.read(artifact(done['source']))==base
  expected=[(block,case) for block in range(plan[lane]['blocks']) for case in (plan['cases'] if block==0 else list(reversed(plan['cases'])))]
  for r,(block,case) in zip(rows,expected):
   assert r['lane']==lane and r['block']==block and r['shape']==case['shape'] and r['mode']==case['mode'] and r['binary']==binaries[index]
   assert r['exit_code']==0 and last<=r['started']<=r['ended'];last=r['ended'];artifact(r['report']);artifact(r['log']);reports+=1
 assert reports==45
 for script in ['analyze.py','root_audit.py','failure_audit.py']:
  subprocess.run([sys.executable,'-B',str(p/script),'--check'],check=True)
 analysis=c.read(p/'analysis.json');audit=c.read(p/'root-audit.json')
 for row in audit['rows']:
  other=analysis['census']['by_case_mode'][row['shape']+'/'+row['mode']]['repeats'][row['block']]
  for left,right in [('instances','raw_instance_rows'),('starts','iterator_starts'),('clones','iterator_clones'),('drops','iterator_drops'),('attribute_count_histogram','attribute_count_histogram'),('element_name_hex_histogram','element_name_hex_histogram')]:assert row[left]==other[right]
  for left,right in [('next_calls','next_calls'),('successful_yields','next_ok'),('error_yields','next_err'),('end_yields','next_none')]:assert row[left]==other['next_transitions'][right]
 inheritance=c.read(p/'probe-inheritance.json');marker=inheritance['unchanged_tail_from'].encode()
 tail=(p/'probe-src/src/main.rs').read_bytes().split(marker,1)[1]
 import hashlib
 assert len(tail)==inheritance['bytes'] and hashlib.sha256(tail).hexdigest()==inheritance['sha256']
 assert tail==(c.ROOT/inheritance['inherited_from']).read_bytes().split(marker,1)[1]
 assert not list(p.rglob('__pycache__'))
 if seal:subprocess.run([sys.executable,'-B',str(p/'seal_packet.py')],check=True)
 print({'reports':reports,'production_change':False,'seal_checked':seal})
if __name__=='__main__':
 ap=argparse.ArgumentParser();ap.add_argument('--require-final-seal',action='store_true');validate(ap.parse_args().require_final_seal)
