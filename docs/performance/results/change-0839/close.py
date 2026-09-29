"""Validate final custody, delete only marked roots, and seal owned files."""
import driver as d
import qualify as q
import readers as r
import re,shutil,sys
REPORT='docs/performance/0839-cached-part-scheduling-requalification.md'
INDEXES=[f'docs/performance/{n}.md' for n in ['BASELINE','HOTSPOTS','REPORT','CRUD_COVERAGE','GOAL_AUDIT']]
def verify():
 origin=d.read(d.P/'origin.json')
 for group in ['normative','unrelated']:
  assert all(d.sha(d.ROOT/n)==h for n,h in origin[group].items())
 analysis=d.read(d.P/'analysis.json');decision=analysis['decision'];expected=d.read(d.P/('candidate-source.json' if decision=='adopt' else 'freeze-before.json'))['source']
 assert d.source()==expected
 admission=d.read(d.P/'admission.json')
 for x in admission['files']:q.bind(x)
 assert admission['source']==d.read(d.P/'candidate-source.json')['source']
 for leg in ['before','after']:
  quality=d.read(d.P/f'quality-{leg}.json')
  for x in quality['commands']:q.bind(x);assert d.read(x['path'])['exit_code']==0
  allocator_leg='before-v2' if leg=='before' else 'after'
  for x in d.read(d.P/f'allocator-quality-{allocator_leg}.json')['receipts']:q.bind(x);assert d.read(x['path'])['exit_code']==0
  for suffix in ['', '-observer','-allocator']:
   b=d.read(d.P/f'build-{leg}{suffix}.json');q.bind(b['freeze']);q.bind(b['receipt']);assert d.read(b['receipt']['path'])['exit_code']==0
 # Replay the syscall witness independently of the capture script.
 for row in d.read(d.P/'trace-analysis.json')['rows']:
  q.bind(row['log']);q.bind(row['report'])
  r.validate_report(d.Path(row['report']['path']),row['case'],1,0)
  text=d.Path(row['log']['path']).read_text()
  created=[line for line in text.splitlines() if any(x in line for x in ['clone(','clone3(','<... clone resumed>','<... clone3 resumed>']) and re.search(r'= [1-9][0-9]*\s*$',line)]
  assert len(created)==row['successful_thread_creations']==row['expected']
 commands=[];tests={}
 for folder in sorted((d.P/'commands').iterdir()):
  receipt=d.read(folder/'receipt.json');started=d.read(folder/'started.json');q.bind(receipt['log'])
  assert receipt['argv']==started['argv'] and receipt['started_unix']==started['started_unix'] and receipt['finished_unix']>=receipt['started_unix']
  expected_code={'before-allocator-test':101,'reader-tests':1}.get(folder.name,0)
  assert receipt['exit_code']==expected_code,(folder.name,receipt['exit_code'])
  commands.append(dict(label=folder.name,exit_code=receipt['exit_code']))
  matches=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;', (folder/'output.log').read_text())
  if matches:tests[folder.name]={key:sum(int(x[i]) for x in matches) for i,key in enumerate(['passed','failed','ignored'])}
 return dict(status='pass',decision=decision,commands=commands,tests=tests,reports=sum(analysis['counts'].values()),samples=analysis['samples'],trace_reports=6,trace_samples=6,regression_flags=len(analysis['regression_flags']),benefit_cases=analysis['benefit_cases'],production_changes=3 if decision=='adopt' else 0)
mode=sys.argv[1]
if mode=='clean':
 d.check();closure=verify();removed=[];binaries=[]
 for leg in ['before','after']:
  for suffix in ['', '-observer','-allocator']:
   b=d.read(d.P/f'build-{leg}{suffix}.json')['binary'];assert d.desc(b['path'])==b;binaries.append(b)
 launcher=d.read(d.P/'launcher-quality.json')['binary'];assert d.desc(launcher['path'])==launcher;binaries.append(launcher)
 for root in [d.TARGET,d.SCRATCH]:
  assert d.read(root/'.owner-0839.json')==dict(packet=str(d.P),path=str(root))
  files=[p for p in root.rglob('*') if p.is_file()];removed.append(dict(path=str(root),files=len(files),logical_bytes=sum(p.stat().st_size for p in files)))
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
print(mode+' PASS')
