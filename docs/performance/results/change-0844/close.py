"""Replay evidence and decision; remove only owned roots; seal exact batch."""
import driver as d
import qualify as q
import analyze as a
import observer_check as oc
import crosscheck as cc
import tables
import mechanism
import re,sys,shutil
REPORT='docs/performance/0844-opc-missing-ancestor-staging.md'
INDEXES=[f'docs/performance/{n}.md' for n in ['BASELINE','HOTSPOTS','REPORT','CRUD_COVERAGE','GOAL_AUDIT']]
PRODUCTION=['crates/litchi-opc/src/members.rs']

def verify():
 origin=d.read(d.P/'origin.json')
 for group in ['normative','unrelated']:assert all(d.sha(d.ROOT/n)==h for n,h in origin[group].items())
 for x in d.read(d.P/'admission.json')['files']:q.bind(x)
 decision=d.read(d.P/'decision.json');assert decision['status']=='complete-batch' and decision['decision'] in ['adopt','reject']
 leg='after' if decision['decision']=='adopt' else 'before'
 assert d.source()==d.read(d.P/f'source-{leg}.json')['source']
 assert d.sha(d.ROOT/'Cargo.lock')==d.sha(d.P/'root-Cargo.lock')
 for leg in ['before','after']:
  build=d.read(d.P/f'build-{leg}.json');q.bind(build['freeze'])
  if d.Path(build['binary']['path']).exists():q.bind(build['binary'])
  else:assert build['binary'] in d.read(d.P/'cleanup.json')['verified_binaries'] and not d.TARGET.exists()
  assert oc.verify(leg)==d.read(d.P/f'observer-preflight-{leg}.json')
 result=a.analyze();assert result==d.read(d.P/'analysis.json') and result['reports']==512 and result['samples']==11648
 assert mechanism.verify()==d.read(d.P/'mechanism.json')
 assert cc.verify()==d.read(d.P/'crosscheck.json') and tables.render()==(d.P/'tables.md').read_text()
 assert decision['analysis']==d.desc(d.P/'analysis.json')
 for field in ['regression_flags','tail_review','spread_flags','allocation_flags']:assert decision[field]==len(result[field])
 if decision['decision']=='adopt':assert decision['material_benefit'] and decision['all_flags_reviewed'] and decision['unresolved_common_workload_damage'] is False
 assert d.read(d.P/'quality-after.json')['status']=='pass'
 for name in ['reader-tests','heap-reader-tests']:
  tests=d.read(d.P/(name+'.json'));assert tests['status']=='pass'
  for k in ['reader','test_source','receipt']:q.bind(tests[k])
 commands=[];totals={}
 for folder in sorted((d.P/'commands').iterdir()):
  end=d.read(folder/'receipt.json');start=d.read(folder/'started.json');q.bind(end['log'])
  assert end['argv']==start['argv'] and end['started_unix']==start['started_unix'] and end['finished_unix']>=end['started_unix']
  assert end['exit_code']==0,(folder.name,end['exit_code'])
  commands.append(dict(label=folder.name,exit_code=end['exit_code']))
  rows=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',(folder/'output.log').read_text())
  if rows:totals[folder.name]={k:sum(int(x[i]) for x in rows) for i,k in enumerate(['passed','failed','ignored'])}
 assert totals['before-opc-test']['failed']==totals['after-owner-test']['failed']==0
 assert not list(d.P.rglob('__pycache__'))
 return dict(status='pass',decision=decision['decision'],reports=512,samples=11648,separate_heap_preflight_reports=2,separate_heap_preflight_samples=2,commands=commands,tests=totals,formal_runs_replaced=0)

if __name__=='__main__':
 mode=sys.argv[1]
 if mode=='clean':
  d.check();value=verify();binaries=[d.read(d.P/f'build-{leg}.json')['binary'] for leg in ['before','after']]
  for x in binaries:q.bind(x)
  removed=[]
  for root in [d.TARGET,d.SCRATCH]:
   assert d.read(root/'.owner-0844.json')==dict(packet=str(d.P),path=str(root))
   files=[p for p in root.rglob('*') if p.is_file()];removed.append(dict(path=str(root),files=len(files),logical_bytes=sum(p.stat().st_size for p in files)))
   shutil.rmtree(root);assert not root.exists()
  d.write(d.P/'cleanup.json',dict(status='pass',verified_binaries=binaries,removed=removed));d.write(d.P/'closure.json',value)
 elif mode=='seal':
  assert verify()==d.read(d.P/'closure.json') and not d.TARGET.exists() and not d.SCRATCH.exists()
  names=(PRODUCTION if d.read(d.P/'decision.json')['decision']=='adopt' else [])+INDEXES+[REPORT]+[str(p.relative_to(d.ROOT)) for p in sorted(d.P.rglob('*')) if p.is_file() and p.name!='seal.json']
  d.write(d.P/'seal.json',dict(base=d.read(d.P/'origin.json')['base'],files={n:d.sha(d.ROOT/n) for n in names}))
 elif mode=='verify':
  value=verify()
  if (d.P/'closure.json').exists():assert value==d.read(d.P/'closure.json')
  if (d.P/'seal.json').exists():
   seal=d.read(d.P/'seal.json');assert all(d.sha(d.ROOT/n)==h for n,h in seal['files'].items()) and not d.TARGET.exists() and not d.SCRATCH.exists()
   if d.output(['git','rev-parse','HEAD'])!=seal['base']:
    assert d.output(['git','rev-parse','HEAD^'])==seal['base']
    assert set(d.output(['git','diff-tree','--no-commit-id','--name-only','-r','HEAD']).splitlines())==set(seal['files'])|{str((d.P/'seal.json').relative_to(d.ROOT))}
 else:raise AssertionError(mode)
 print(mode,'PASS')
