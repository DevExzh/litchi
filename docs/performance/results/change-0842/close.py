"""Audit evidence/decision, remove only owned roots, and seal exact commit paths."""
import driver as d
import qualify as q
import analyze as a
import codec_check_v2 as c
import re,shutil,sys
REPORT='docs/performance/0842-opc-runtime-compression-dispatch.md'
INDEXES=[f'docs/performance/{name}.md' for name in ['BASELINE','HOTSPOTS','REPORT','CRUD_COVERAGE','GOAL_AUDIT']]
PRODUCTION=['Cargo.toml','crates/soapberry-zip/Cargo.toml','tools/perf-baseline/Cargo.lock','tools/perf-execution/Cargo.lock','tools/native-resave/Cargo.lock']
FAILURES={'before-probe-clippy':101,'qualification-00-tiny-write-leg-a':1,'baseline-qualification-reader':1,'after-probe-clippy':101,'after-workspace-check':101,'baseline-iwork-example-control':101,'after-final-workspace-doc':101,'baseline-iwork-doc-control':101}

def verify():
 origin=d.read(d.P/'origin.json')
 for group in ['normative','unrelated']:
  assert all(d.sha(d.ROOT/n)==h for n,h in origin[group].items())
 for item in d.read(d.P/'admission.json')['files']:q.bind(item)
 decision=d.read(d.P/'decision.json');assert decision['status']=='complete-batch' and decision['decision'] in ['adopt','reject']
 expected=d.read(d.P/('source-final.json' if decision['decision']=='adopt' else 'source-before.json'))['source']
 assert d.source()==expected
 assert d.sha(d.ROOT/'Cargo.lock')==d.sha(d.P/'source-before/Cargo.lock')
 for leg in ['before','after']:
  build=d.read(d.P/f'build-{leg}.json');q.bind(build['freeze'])
  for key in ['binary','codec']:
   binary=build[key]
   if d.Path(binary['path']).is_file():q.bind(binary)
   else:assert binary in d.read(d.P/'cleanup.json')['verified_binaries'] and not d.TARGET.exists()
 analysis=a.analyze();assert analysis==d.read(d.P/'analysis.json')
 assert analysis['reports']==168 and analysis['samples']==4344
 assert [x['row'] for x in sorted(analysis['native'],key=lambda x:x['started'])]==d.read(d.P/'plan.json')['native']
 assert c.verify()==d.read(d.P/'codecs-validation-v2.json')
 assert decision['analysis']==d.desc(d.P/'analysis.json')
 assert decision['regression_flags']==len(analysis['regression_flags']) and decision['spread_flags']==len(analysis['spread_flags']) and decision['tail_review_points']==len(analysis['tail_review'])
 if decision['decision']=='adopt':assert decision['material_benefit'] and decision['all_flags_reviewed'] and decision['unresolved_common_workload_damage'] is False
 check=d.read(d.P/'crosscheck.json');assert check['status']=='pass' and check['quantile_rows']==36 and check['metric_rows']==84;q.bind(check['analysis']);q.bind(check['checker'])
 tests=d.read(d.P/'reader-tests.json');assert len(tests['checks'])==19 and tests['status']=='pass'
 for key in ['reader','tests','receipt']:q.bind(tests[key])
 assert d.read(d.P/'quality-after.json')['status']=='pass'
 assert d.read(d.P/'tool-checks.json')['status']=='pass'
 for item in d.read(d.P/'tool-checks.json')['receipts']:assert d.read(q.bind(item))['exit_code']==0
 for item in d.read(d.P/'quality-after.json')['receipts']:assert d.read(q.bind(item))['exit_code']==0
 commands=[];totals={}
 for folder in sorted((d.P/'commands').iterdir()):
  end=d.read(folder/'receipt.json');start=d.read(folder/'started.json');q.bind(end['log'])
  assert end['argv']==start['argv'] and end['started_unix']==start['started_unix'] and end['finished_unix']>=end['started_unix']
  assert end['exit_code']==FAILURES.get(folder.name,0),(folder.name,end['exit_code'])
  commands.append(dict(label=folder.name,exit_code=end['exit_code']))
  rows=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',(folder/'output.log').read_text())
  if rows:totals[folder.name]={k:sum(int(x[i]) for x in rows) for i,k in enumerate(['passed','failed','ignored'])}
 assert totals['before-opc-test']==dict(passed=1650,failed=0,ignored=3)
 assert totals['after-final-owner-test']['failed']==0
 failed=d.P/'runs/qualification-00-tiny-write-leg-a';receipt=d.read(failed/'receipt.json');assert receipt['exit_code']==1
 for item in receipt['artifacts']:q.bind(item)
 q.bind(receipt['command']);assert not (failed/'report.json').exists()
 old=d.read(failed/'started.json')['build'];assert d.sha(d.P/'build-before-initial.json')==old['sha256']
 return dict(status='pass',decision=decision['decision'],reports=168,samples=4344,codec_processes=2,codec_archives=18,failed_preflight_processes=1,formal_native_runs_replaced=0,regression_flags=len(analysis['regression_flags']),spread_flags=len(analysis['spread_flags']),commands=commands,tests=totals)

mode=sys.argv[1]
if mode=='clean':
 d.check();closure=verify();binaries=[]
 for leg in ['before','after']:
  build=d.read(d.P/f'build-{leg}.json')
  for key in ['binary','codec']:q.bind(build[key]);binaries.append(build[key])
 old=d.read(d.P/'build-before-initial.json');initial=[]
 for key in ['binary','codec']:
  path=d.TARGET/'retained/initial-leg-a'/d.Path(old[key]['path']).name
  actual=d.desc(path);assert actual['sha256']==old[key]['sha256'] and actual['bytes']==old[key]['bytes'];initial.append(actual)
 removed=[]
 for root in [d.TARGET,d.SCRATCH]:
  assert d.read(root/'.owner-0842.json')==dict(packet=str(d.P),path=str(root))
  files=[x for x in root.rglob('*') if x.is_file()];removed.append(dict(path=str(root),files=len(files),logical_bytes=sum(x.stat().st_size for x in files)))
  shutil.rmtree(root);assert not root.exists()
 d.write(d.P/'cleanup.json',dict(status='pass',verified_binaries=binaries,verified_initial_binaries=initial,removed=removed));d.write(d.P/'closure.json',closure)
elif mode=='seal':
 assert verify()==d.read(d.P/'closure.json');assert not d.TARGET.exists() and not d.SCRATCH.exists()
 production=PRODUCTION if d.read(d.P/'decision.json')['decision']=='adopt' else []
 names=production+INDEXES+[REPORT]+[str(p.relative_to(d.ROOT)) for p in sorted(d.P.rglob('*')) if p.is_file() and p.name!='seal.json']
 d.write(d.P/'seal.json',dict(base=d.read(d.P/'origin.json')['base'],files={n:d.sha(d.ROOT/n) for n in names}))
elif mode=='verify':
 closure=verify()
 if (d.P/'closure.json').exists():assert closure==d.read(d.P/'closure.json')
 if (d.P/'seal.json').exists():
  seal=d.read(d.P/'seal.json');assert all(d.sha(d.ROOT/n)==h for n,h in seal['files'].items());assert not d.TARGET.exists() and not d.SCRATCH.exists()
  if d.output(['git','rev-parse','HEAD'])!=seal['base']:
   assert d.output(['git','rev-parse','HEAD^'])==seal['base']
   changed=set(d.output(['git','diff-tree','--no-commit-id','--name-only','-r','HEAD']).splitlines());assert changed==set(seal['files'])|{str((d.P/'seal.json').relative_to(d.ROOT))}
else:raise AssertionError(mode)
print(mode,'PASS')
