"""Final direct-helper packet custody; analysis separately validates raw reports."""
import argparse,subprocess,sys
from pathlib import Path
import custody as c
from fixture_check import check

def artifact(a):
 p=Path(a['path']);assert p.is_file() and not p.is_symlink() and c.artifact(p)==a,p
 return p

def validate(seal=False):
 p=c.P;origin=c.read(p/'origin.json');source=c.read(p/'source.json');plan=c.read(p/'plan.json')
 assert source['revision']==origin['base'] and c.source()['files']==source['files']
 worktrees=subprocess.check_output(['git','worktree','list','--porcelain'],cwd=c.ROOT,text=True)
 other=lambda value:[block for block in value.strip().split('\n\n') if not block.startswith('worktree '+str(c.ROOT)+'\n')]
 assert other(worktrees)==other(origin['worktrees'])
 for n,h in origin['unrelated'].items():assert c.sha(c.ROOT/n)==h,n
 for n,h in c.read(p/'architecture-inputs.json').items():assert c.sha(c.ROOT/n)==h,n
 assert c.sha(p/'workspace-Cargo.lock')==c.sha(c.ROOT/'Cargo.lock')
 old=c.read(p/'inheritance.json')
 for n in ['baseline_helper','candidate_helper','previous_seal','candidate_manifest','parser']:artifact(old[n])
 assert c.sha(p/'callgrind_parser.py')==old['parser']['sha256']
 assert c.sha(p/'probe-src/src/baseline.rs')==old['baseline_helper']['sha256']
 assert c.sha(p/'probe-src/src/candidate.rs')==old['candidate_helper']['sha256']
 b=c.read(p/'build/build.json');inputs=c.read(artifact(b['inputs']))
 for n,h in inputs['frozen'].items():assert c.sha(p/n)==h,n
 for n,h in inputs['probe'].items():assert c.sha(p/'probe-src'/n)==h,n
 artifact(b['lock']);assert len(b['rows'])==6
 assert b['environment']=={'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','RUSTFLAGS':None}
 last=0
 q=c.read(p/'quality/complete.json');assert len(q['rows'])==6
 for n,h in c.read(p/'quality/archive-inputs.json').items():assert c.sha(p/'candidate'/n)==h
 for r in q['rows']:
  assert r['exit_code']==0 and last<=r['started']<=r['ended'];last=r['ended'];artifact(r['log'])
 for leg,files in q['inputs'].items():
  for n,h in files.items():assert c.sha(p/'test-src'/leg/n)==h
  for owner in ['litchi-opc','litchi-ole-common','litchi-sign','litchi-xldm','xml-minifier']:
   assert c.sha(p/'test-src'/leg/'crates'/owner/'src/xml_attributes.rs')==c.sha(p/'candidate'/leg/(owner+'-xml_attributes.rs'))
  assert c.sha(p/'test-src'/leg/'crates/litchi-opc/src/xml_attributes/tests.rs')==c.sha(p/'candidate'/leg/'litchi-opc-xml_attributes-tests.rs')
 for r in b['rows']:
  assert r['exit_code']==0 and last<=r['started']<=r['ended'];last=r['ended'];artifact(r['log'])
  if 'output' in r:artifact(r['output'])
 binary=b['binary']
 if Path(binary['path']).exists():artifact(binary)
 else:
  cleanup=c.read(p/'cleanup.json');assert cleanup['removed_binaries']==[binary] and cleanup['target_removed'] and cleanup['target']==str(c.TARGET) and not c.TARGET.exists()
 check();cases=c.read(p/'cases.json');assert len(cases)==33;ids=[r['id'] for r in cases];assert len(set(ids))==33
 reports=0;samples=0
 for lane,count in [('native',792),('profiles',264)]:
  done=c.read(p/lane/'complete.json');rows=c.read(artifact(done['receipts']));assert len(rows)==done['children']==count
  assert c.read(artifact(done['source']))==source
  schedule=[(block,case,mode,leg) for block,order in enumerate(plan[lane]['orders']) for case in ids for mode in ['construct','consume'] for leg in order]
  for r,expected in zip(rows,schedule):
   assert (r['block'],r['case'],r['mode'],r['leg'])==expected
   assert r['exit_code']==0 and r['binary']==binary and last<=r['started']<=r['ended'];last=r['ended']
   artifact(r['report']);artifact(r['log'])
   if lane=='profiles':
    for a in r['artifacts'].values():artifact(a)
   reports+=1;samples+=plan[lane]['samples']
 assert (reports,samples)==(1056,24024)
 for script in ['failure_audit.py','analyze.py','root_cg_totals.py','root_native_audit.py','decision.py']:
  subprocess.run([sys.executable,'-B',str(p/script),'--check'],check=True)
 a=c.read(p/'analysis.json')['native']['analysis']['paired_by_case_mode']
 for row in c.read(p/'root-native-audit.json')['rows']:
  other=a[row['case']+'/'+row['mode']]
  assert row['ratio']==other['ratio_median'] and row['ci_low']==other['bootstrap']['ci_low'] and row['ci_high']==other['bootstrap']['ci_high'] and row['regression']==other['diagnostic_regression']
 assert not list(p.rglob('__pycache__'))
 if seal:subprocess.run([sys.executable,'-B',str(p/'seal_packet.py')],check=True)
 print({'reports':reports,'samples':samples,'production_change':False,'seal_checked':seal})

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--require-final-seal',action='store_true');args=parser.parse_args();validate(args.require_final_seal)
