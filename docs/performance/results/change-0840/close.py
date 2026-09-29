"""Verify final source/evidence, remove only owned roots, seal exact commit paths."""
import driver as d
import qualify as q
import re,shutil,sys
REPORT='docs/performance/0840-cfb-fresh-emission-order.md'
INDEXES=[f'docs/performance/{name}.md' for name in ['BASELINE','HOTSPOTS','REPORT','CRUD_COVERAGE','GOAL_AUDIT']]
FAILURES={'before-probe-clippy':101,'qualify-before':1,'qualify-before-v2':1,'focused-red':101,'after-fmt':1,'paired-analysis':1}

def verify():
 origin=d.read(d.P/'origin.json')
 for group in ['normative','unrelated']:
  assert all(d.sha(d.ROOT/n)==h for n,h in origin[group].items())
 analysis=d.read(d.P/'analysis.json');q.bind(analysis['analyzer']);q.bind(d.read(d.P/'analysis-repair.json')['corrected']);adopt=analysis['decision']=='adopt'
 assert d.source()==d.read(d.P/('candidate-source.json' if adopt else 'source-before.json'))['source']
 for x in d.read(d.P/'admission.json')['files']:q.bind(x)
 for leg in ['before','after']:
  for name in [f'quality-{leg}.json',f'probe-quality-{leg}.json']:
   value=d.read(d.P/name);assert value['status']=='pass'
   for x in value['receipts']:q.bind(x);assert d.read(x['path'])['exit_code']==0
  for kind in ['native','allocation']:
   b=d.read(d.P/f'build-{leg}-{kind}.json');q.bind(b['freeze']);q.bind(b['receipt']);assert d.read(b['receipt']['path'])['exit_code']==0
 commands=[];tests={};build_intervals=[]
 for folder in sorted((d.P/'commands').iterdir()):
  receipt=d.read(folder/'receipt.json');started=d.read(folder/'started.json');q.bind(receipt['log'])
  assert receipt['argv']==started['argv'] and receipt['started_unix']==started['started_unix'] and receipt['finished_unix']>=receipt['started_unix']
  assert receipt['exit_code']==FAILURES.get(folder.name,0),(folder.name,receipt['exit_code'])
  commands.append(dict(label=folder.name,exit_code=receipt['exit_code']))
  if receipt['argv'][0]=='cargo':build_intervals.append((receipt['started_unix'],receipt['finished_unix']))
  matches=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;', (folder/'output.log').read_text())
  if matches:tests[folder.name]={key:sum(int(x[i]) for x in matches) for i,key in enumerate(['passed','failed','ignored'])}
 # Formal workloads must not overlap any root Cargo invocation.
 for lane in ['native','allocation']:
  for path in (d.P/'runs'/lane).glob('*/started.json'):
   start=d.read(path);end=d.read(path.parent/'receipt.json')['finished_unix']
   assert all(end<=a or start['started_unix']>=b for a,b in build_intervals)
 assert analysis['reports']==270 and analysis['samples']==6534
 return dict(status='pass',decision=analysis['decision'],commands=commands,tests=tests,reports=analysis['reports'],samples=analysis['samples'],regression_flags=len(analysis['regression_flags']),target_benefit=analysis['target_benefit'],production_changes=2 if adopt else 0)

mode=sys.argv[1]
if mode=='clean':
 d.check();closure=verify();binaries=[]
 for leg in ['before','after']:
  for kind in ['native','allocation']:
   b=d.read(d.P/f'build-{leg}-{kind}.json')['binary'];assert d.desc(b['path'])==b;binaries.append(b)
 removed=[]
 for root in [d.TARGET,d.SCRATCH]:
  assert d.read(root/'.owner-0840.json')==dict(packet=str(d.P),path=str(root))
  files=[p for p in root.rglob('*') if p.is_file()]
  removed.append(dict(path=str(root),files=len(files),logical_bytes=sum(p.stat().st_size for p in files)))
  shutil.rmtree(root);assert not root.exists()
 d.write(d.P/'cleanup.json',dict(status='pass',removed=removed,verified_binaries=binaries))
 d.write(d.P/'closure.json',closure)
elif mode=='seal':
 assert verify()==d.read(d.P/'closure.json');assert not d.TARGET.exists() and not d.SCRATCH.exists()
 names=INDEXES+[REPORT]+[str(p.relative_to(d.ROOT)) for p in sorted(d.P.rglob('*')) if p.is_file() and p.name!='seal.json']
 if d.read(d.P/'analysis.json')['decision']=='adopt':names+=d.read(d.P/'candidate-source.json')['changed']
 d.write(d.P/'seal.json',dict(base=d.read(d.P/'origin.json')['base'],files={n:d.sha(d.ROOT/n) for n in names}))
elif mode=='verify':
 assert verify()==d.read(d.P/'closure.json');seal=d.read(d.P/'seal.json')
 assert all(d.sha(d.ROOT/n)==h for n,h in seal['files'].items())
 assert not d.TARGET.exists() and not d.SCRATCH.exists()
 if d.output(['git','rev-parse','HEAD'])!=seal['base']:
  assert d.output(['git','rev-parse','HEAD^'])==seal['base']
  changed=set(d.output(['git','diff-tree','--no-commit-id','--name-only','-r','HEAD']).splitlines())
  assert changed==set(seal['files'])|{str((d.P/'seal.json').relative_to(d.ROOT))}
else:raise AssertionError(mode)
print(mode,'PASS')
