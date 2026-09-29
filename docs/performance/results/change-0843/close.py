"""Replay retained evidence, clean only owned roots, seal exact commit paths."""
import driver as d
import qualify as q
import analyze as a
import archive_check_v2 as ac
import observer_check as oc
import crosscheck as cc
import tables
import sys,shutil
REPORT='docs/performance/0843-opc-many-part-publication-profile.md'
INDEXES=[f'docs/performance/{n}.md' for n in ['BASELINE','HOTSPOTS','REPORT','CRUD_COVERAGE','GOAL_AUDIT']]

def verify():
 origin=d.read(d.P/'origin.json')
 for group in ['normative','unrelated']:
  assert all(d.sha(d.ROOT/n)==h for n,h in origin[group].items())
 assert d.source()==d.read(d.P/'source.json')['source']
 assert d.sha(d.ROOT/'Cargo.lock')==d.sha(d.P/'root-Cargo.lock')
 for x in d.read(d.P/'admission.json')['files']:q.bind(x)
 build=d.read(d.P/'build.json');q.bind(build['freeze'])
 if d.Path(build['binary']['path']).exists():q.bind(build['binary'])
 else:assert build['binary']==d.read(d.P/'cleanup.json')['verified_binary'] and not d.TARGET.exists()
 result=a.analyze();assert result==d.read(d.P/'analysis.json') and result['reports']==72 and result['samples']==1464
 for lane in ['native','profiles','heap']:
  assert [x['row'] for x in sorted(result['populations'][lane],key=lambda x:x['started'])]==d.read(d.P/'plan.json')[lane]
 assert ac.verify()==d.read(d.P/'archive-check-v2.json')
 assert oc.verify()==d.read(d.P/'observer-preflight-validation.json')
 assert cc.verify()==d.read(d.P/'crosscheck.json')
 assert tables.render()==(d.P/'tables.md').read_text()
 decision=d.read(d.P/'decision.json');assert decision['decision']=='diagnostic-only' and decision['analysis']==d.desc(d.P/'analysis.json')
 assert decision['spread_flags']==len(result['spread_flags']) and decision['production_changed'] is False
 for name in ['reader-tests','heap-reader-tests']:
  tests=d.read(d.P/(name+'.json'));assert tests['status']=='pass'
  for k in ['reader','test_source','receipt']:q.bind(tests[k])
 commands=[]
 for folder in sorted((d.P/'commands').iterdir()):
  end=d.read(folder/'receipt.json');start=d.read(folder/'started.json');q.bind(end['log'])
  assert end['argv']==start['argv'] and end['started_unix']==start['started_unix'] and end['finished_unix']>=end['started_unix']
  assert end['exit_code']==(101 if folder.name in ['probe-clippy','probe-v2-clippy'] else 0),(folder.name,end['exit_code'])
  commands.append(dict(label=folder.name,exit_code=end['exit_code']))
 assert not list(d.P.rglob('__pycache__'))
 return dict(status='pass',reports=72,samples=1464,separate_observer_preflight_reports=2,separate_observer_preflight_samples=2,production_changed=False,spread_flags=len(result['spread_flags']),commands=commands)

if __name__=='__main__':
 mode=sys.argv[1]
 if mode=='clean':
  d.check();value=verify();binary=d.read(d.P/'build.json')['binary'];q.bind(binary);removed=[]
  for root in [d.TARGET,d.SCRATCH]:
   assert d.read(root/'.owner-0843.json')==dict(packet=str(d.P),path=str(root))
   files=[p for p in root.rglob('*') if p.is_file()];removed.append(dict(path=str(root),files=len(files),logical_bytes=sum(p.stat().st_size for p in files)))
   shutil.rmtree(root);assert not root.exists()
  d.write(d.P/'cleanup.json',dict(status='pass',verified_binary=binary,removed=removed));d.write(d.P/'closure.json',value)
 elif mode=='seal':
  assert verify()==d.read(d.P/'closure.json') and not d.TARGET.exists() and not d.SCRATCH.exists()
  names=INDEXES+[REPORT]+[str(p.relative_to(d.ROOT)) for p in sorted(d.P.rglob('*')) if p.is_file() and p.name!='seal.json']
  d.write(d.P/'seal.json',dict(base=d.read(d.P/'origin.json')['base'],files={n:d.sha(d.ROOT/n) for n in names}))
 elif mode=='verify':
  value=verify()
  if (d.P/'closure.json').exists():assert value==d.read(d.P/'closure.json')
  if (d.P/'seal.json').exists():
   seal=d.read(d.P/'seal.json');assert all(d.sha(d.ROOT/n)==h for n,h in seal['files'].items())
   assert not d.TARGET.exists() and not d.SCRATCH.exists()
   if d.output(['git','rev-parse','HEAD'])!=seal['base']:
    assert d.output(['git','rev-parse','HEAD^'])==seal['base']
    assert set(d.output(['git','diff-tree','--no-commit-id','--name-only','-r','HEAD']).splitlines())==set(seal['files'])|{str((d.P/'seal.json').relative_to(d.ROOT))}
 else:raise AssertionError(mode)
 print(mode,'PASS')
