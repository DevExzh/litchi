"""Verify restored source and evidence, remove owned roots, seal exact commit paths."""
import driver as d
import qualify_v3 as q
import analyze_v3 as a
import re,shutil,sys
REPORT='docs/performance/0841-ppt-borrowed-text-instruction-diagnostic.md'
INDEXES=[f'docs/performance/{name}.md' for name in ['BASELINE','HOTSPOTS','REPORT','CRUD_COVERAGE','GOAL_AUDIT']]
FAILURES={'before-probe-clippy':101,'before-v2-probe-test-all':101,'paired-analysis':1}

def verify():
 origin=d.read(d.P/'origin.json')
 for group in ['normative','unrelated']:
  assert all(d.sha(d.ROOT/n)==h for n,h in origin[group].items())
 baseline=d.read(d.P/'source-before.json')['source']
 assert d.source()==baseline==d.read(d.P/'restored.json')['source']
 for name in ['admission.json','admission-v2.json']:
  for x in d.read(d.P/name)['files']:q.bind(x)
 repair=d.read(d.P/'analysis-repair.json')
 for key in ['original_qualifier','corrected_qualifier','original_analyzer','corrected_analyzer','failure','success']:q.bind(repair[key])
 assert repair['workload_runs_replaced']==0 and not repair['metrics_changed'] and not repair['thresholds_changed']
 assert repair['corrected_qualifier']==d.desc(d.P/'qualify_v3.py') and repair['corrected_analyzer']==d.desc(d.P/'analyze_v3.py')
 reader_repair=d.read(d.P/'reader-repair.json');q.bind(reader_repair['corrected_reader']);q.bind(reader_repair['original_reader']);q.bind(reader_repair['initial_replacement_archive']);q.bind(reader_repair['violation'])
 violation=d.read(d.P/'reader-freeze-violation.json')
 assert violation['old']['sha256']==d.sha(d.P/'readers.py') and violation['new']['sha256']==d.sha(d.P/'readers_v2_initial.py')
 for leg in ['before','after']:
  for name in [f'quality-{leg}.json',f'probe-quality-{leg}.json']:
   value=d.read(d.P/name);assert value['status']=='pass'
   for x in value['receipts']:assert d.read(q.bind(x))['exit_code']==0
  build=d.read(d.P/f'build-{leg}.json');q.bind(build['freeze']);assert d.read(q.bind(build['receipt']))['exit_code']==0
  binary=build['binary']
  if d.Path(binary['path']).is_file():q.bind(binary)
  else:
   cleanup=d.read(d.P/'cleanup.json');assert binary in cleanup['verified_binaries'] and not d.TARGET.exists()
 analysis=a.analyze();assert analysis==d.read(d.P/'analysis.json')
 assert analysis['reports']==156 and analysis['samples']==3636 and analysis['decision']=='diagnostic-only-restore-production'
 tests_final=d.read(d.P/'reader-tests-final.json');assert tests_final['status']=='pass' and tests_final['mutation_count']==21
 for key in ['reader','test_implementation','test_runner']:q.bind(tests_final[key])
 for name in ['arithmetic-crosscheck.json','interval-crosscheck.json']:
  value=d.read(d.P/name);assert value['status']=='pass';q.bind(value['analysis'])
 assert d.read(d.P/'arithmetic-crosscheck.json')['native_quantile_rows']==30
 assert d.read(d.P/'interval-crosscheck.json')['paired_intervals']==60
 commands=[];tests={}
 for folder in sorted((d.P/'commands').iterdir()):
  end=d.read(folder/'receipt.json');start=d.read(folder/'started.json');q.bind(end['log'])
  assert end['argv']==start['argv'] and end['started_unix']==start['started_unix'] and end['finished_unix']>=end['started_unix']
  assert end['exit_code']==FAILURES.get(folder.name,0),(folder.name,end['exit_code'])
  commands.append(dict(label=folder.name,exit_code=end['exit_code']))
  rows=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',(folder/'output.log').read_text())
  if rows:tests[folder.name]={k:sum(int(x[i]) for x in rows) for i,k in enumerate(['passed','failed','ignored'])}
 assert tests['before-test']==dict(passed=1284,failed=0,ignored=11)
 assert tests['after-test']==dict(passed=1287,failed=0,ignored=11)
 return dict(status='pass',decision=analysis['decision'],reports=156,samples=3636,regression_flags=len(analysis['regression_flags']),spread_flags=len(analysis['spread_flags']),commands=commands,tests=tests,production_changes=0,workload_runs_replaced=0)

mode=sys.argv[1]
if mode=='clean':
 d.check();closure=verify();binaries=[]
 for leg in ['before','after']:
  binary=d.read(d.P/f'build-{leg}.json')['binary'];q.bind(binary);binaries.append(binary)
 removed=[]
 for root in [d.TARGET,d.SCRATCH]:
  assert d.read(root/'.owner-0841.json')==dict(packet=str(d.P),path=str(root))
  files=[p for p in root.rglob('*') if p.is_file()]
  removed.append(dict(path=str(root),files=len(files),logical_bytes=sum(p.stat().st_size for p in files)))
  shutil.rmtree(root);assert not root.exists()
 d.write(d.P/'cleanup.json',dict(status='pass',removed=removed,verified_binaries=binaries))
 d.write(d.P/'closure.json',closure)
elif mode=='seal':
 assert verify()==d.read(d.P/'closure.json');assert not d.TARGET.exists() and not d.SCRATCH.exists()
 names=INDEXES+[REPORT]+[str(p.relative_to(d.ROOT)) for p in sorted(d.P.rglob('*')) if p.is_file() and p.name!='seal.json']
 d.write(d.P/'seal.json',dict(base=d.read(d.P/'origin.json')['base'],files={n:d.sha(d.ROOT/n) for n in names}))
elif mode=='verify':
 closure=verify()
 if (d.P/'closure.json').exists():assert closure==d.read(d.P/'closure.json')
 if (d.P/'seal.json').exists():
  seal=d.read(d.P/'seal.json');assert all(d.sha(d.ROOT/n)==h for n,h in seal['files'].items())
  assert not d.TARGET.exists() and not d.SCRATCH.exists()
  if d.output(['git','rev-parse','HEAD'])!=seal['base']:
   assert d.output(['git','rev-parse','HEAD^'])==seal['base']
   changed=set(d.output(['git','diff-tree','--no-commit-id','--name-only','-r','HEAD']).splitlines())
   assert changed==set(seal['files'])|{str((d.P/'seal.json').relative_to(d.ROOT))}
else:raise AssertionError(mode)
print(mode,'PASS')
