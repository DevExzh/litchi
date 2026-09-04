#!/usr/bin/env python3
"""Corrected XLSX compressed-member attribution; descriptive evidence only."""
import datetime, hashlib, json, os, shutil, subprocess, sys, time
from pathlib import Path
if not __debug__: raise RuntimeError('Assertions must be enabled')
repo=Path.cwd(); out=Path('/tmp/litchi-goal-0409-corrected'); out.mkdir(exist_ok=True)
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def write(n,v): (out/n).write_text(json.dumps(v,indent=2)+'\n')
def git(*args): return subprocess.check_output(['git',*args],text=True).strip()
def state(): return {'revision':git('rev-parse','HEAD'),'status':git('status','--porcelain=v1'),'goal_sha256':sha(repo/'docs/GOAL.md')}
before=state(); assert before['status']=='?? docs/GOAL.md'
if (out/'environment.json').exists(): assert json.loads((out/'environment.json').read_text())['source_before']==before
binary=Path('/tmp/litchi-goal-0409-target/release/litchi-perf-baseline'); identity={'path':str(binary),'bytes':binary.stat().st_size,'sha256':sha(binary)}
flags='-C force-frame-pointers=yes -C force-unwind-tables=yes'
env=os.environ.copy(); env.update(RUSTUP_TOOLCHAIN='1.98.1',RUSTFLAGS=flags,DEBUGINFOD_URLS='')
commands=json.loads((out/'commands.json').read_text()) if (out/'commands.json').exists() else []
def run(label,argv):
 previous=[c for c in commands if c['label']==label]
 if previous:
  assert len(previous)==1 and previous[0]['exit_code']==0
  assert previous[0]['argv']==[str(x) for x in argv]
  return
 t=time.monotonic(); utc=datetime.datetime.now(datetime.timezone.utc).isoformat()
 with (out/(label+'.stdout')).open('wb') as stdout, (out/(label+'.stderr')).open('wb') as stderr:
  p=subprocess.run([str(x) for x in argv],env=env,stdout=stdout,stderr=stderr)
 commands.append({'label':label,'argv':[str(x) for x in argv],'started_utc':utc,'wall_seconds':time.monotonic()-t,'exit_code':p.returncode});write('commands.json',commands)
 if p.returncode: raise RuntimeError(f'{label} failed')
shutil.copy2(__file__,out/'capture.py')
write('environment.json',{'source_before':before,'binary':identity,'runtime_environment':{'RUSTFLAGS':flags,'RUSTUP_TOOLCHAIN':'1.98.1'},'build_environment':{'CARGO_TARGET_DIR':'/tmp/litchi-goal-0409-target','CARGO_BUILD_JOBS':'4','CARGO_INCREMENTAL':'0','CARGO_PROFILE_RELEASE_DEBUG':'1','RUSTFLAGS':flags},'cpu':2,'performance_claim':'none','note':'Corrected observer performs compressed member range intersections during timed source calls; old/new latency comparison is not authorized.'})
run('normal-build-id',['readelf','-n',binary])
sys.path.insert(0,str(repo)); from tools import perf_compare, validate_perf_corpus_binding
specs=[('source-edit','xlsx_source_backed_cell_values_one_edit_save',500,20,1),('multi-sheet-smoke','xlsx_source_backed_cell_values_multi_sheet_edit_save',3,1,2),('all-sheet-batch-smoke','xlsx_source_backed_cell_values_batch_edit_save',3,1,4),('managed-edit-smoke','xlsx_source_backed_managed_cell_values_one_edit_save',3,1,1)]
validation=[]
for label,case,samples,warmup,sheets in specs:
 argv=[binary,'--case',case,'--filesystem-cache','warm','--xlsx-cell-crud-shape','medium','--samples',samples,'--warmup',warmup,'--json',out/(label+'.json'),'--corpus-manifest',out/(label+'.catalog.json')]
 run(label,['taskset','-c','2','/usr/bin/time','-v','-o',out/(label+'.time.txt'),*argv])
 d=json.loads((out/(label+'.json')).read_text()); validate_perf_corpus_binding.validate_paths(out/(label+'.json'),out/(label+'.catalog.json'));perf_compare.validate_parallel_metrics(d)
 assert d['environment']['git_revision']==before['revision'];assert d['environment']['git_worktree_dirty'] is True
 assert d['binary_identity']['binary_sha256']==identity['sha256'];assert d['environment']['cpu_affinity']=='2';assert d['environment']['rustflags']==flags
 assert d['configuration']['xlsx_cell_values_range_accounting']==perf_compare.XLSX_CELL_VALUES_RANGE_ACCOUNTING_VERSION
 perf_compare._validate_xlsx_cell_values_range_accounting(d['configuration'],label)
 r=d['results'][0];assert r['case']==case;assert len(r['elapsed_ns']['samples'])==samples
 assert r['corpus']['archive_sha256']=='dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036'
 m=r['source']['xlsx_cell_values'];assert m['selected_worksheet_count']==sheets
 for k in ['workbook_read_calls','workbook_read_bytes','selected_worksheet_read_calls','selected_worksheet_read_bytes']:
  assert len(m[k])==samples and all(x>0 for x in m[k]),k
 for k in ['unselected_worksheet_read_calls','unselected_worksheet_read_bytes']:
  assert all(x==0 for x in m[k]) if sheets==4 else all(x>0 for x in m[k]),k
 for sorted_index,raw_index in enumerate(r['elapsed_ns']['sample_order']):
  assert sum(m[k][raw_index] for k in ['open_ns','plan_ns','commit_ns','publication_ns'])==r['elapsed_ns']['samples'][sorted_index]
 if label=='source-edit':
  for k,v in [('source_read_calls',257),('source_read_bytes',4233005),('output_sha256','9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1'),('semantic_sha256','3cd21160d4f74fa0f097ab40be08e211b3e460cea788aa2b6705a55fdece07de'),('untouched_member_sha256','7105fcbce160328f666e69fcfd18da9e19fd71dd7b63961e7cddd29d5da1a17d')]: assert m[k]==[v]*samples,k
 if label=='managed-edit-smoke': assert m['output_budget_refusal']['zero_output_verified'] is True
 validation.append({'case':case,'samples':samples,'warmup':warmup,'selected_worksheet_count':sheets,'status':'passed','compressed_overlap':{k:sorted(set(m[k])) for k in ['workbook_read_calls','workbook_read_bytes','selected_worksheet_read_calls','selected_worksheet_read_bytes','unselected_worksheet_read_calls','unselected_worksheet_read_bytes']}})
assert state()==before;assert sha(binary)==identity['sha256']
write('validation.json',{'source_after':state(),'reports':validation,'performance_claim':'none'})
print('Corrected capture validated:',out)
