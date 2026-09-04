#!/usr/bin/env python3
"""Current-state XLSX semantic baseline; no before/after claim."""
import datetime, hashlib, json, os, shutil, subprocess, sys, time
from pathlib import Path
if not __debug__: raise RuntimeError('Assertions must be enabled')
repo=Path.cwd();out=Path('/tmp/litchi-goal-0409-capture');out.mkdir()
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def write(n,d): (out/n).write_text(json.dumps(d,indent=2)+'\n')
def git(*a):return subprocess.check_output(['git',*a],text=True).strip()
def state():return {'revision':git('rev-parse','HEAD'),'status':git('status','--porcelain=v1'),'goal_sha256':sha(repo/'docs/GOAL.md')}
before=state();assert before['status']=='?? docs/GOAL.md';assert before['revision']=='abe38a9570129c6646bb1b1d7207c407fc86c3d6'
flags='-C force-frame-pointers=yes -C force-unwind-tables=yes'
env=os.environ.copy();env.update(RUSTUP_TOOLCHAIN='1.98.1',RUSTFLAGS=flags,DEBUGINFOD_URLS='')
bindir=Path('/tmp/litchi-goal-0409-target/release');bins={n:bindir/n for n in ['litchi-perf-baseline','litchi-perf-baseline-alloc']}
identity={n:{'path':str(p),'sha256':sha(p),'bytes':p.stat().st_size} for n,p in bins.items()}
commands=[]
def run(label,cmd):
 t=time.monotonic();utc=datetime.datetime.now(datetime.timezone.utc).isoformat()
 with (out/(label+'.stdout')).open('wb') as stdout,(out/(label+'.stderr')).open('wb') as stderr:
  r=subprocess.run([str(x) for x in cmd],env=env,stdout=stdout,stderr=stderr)
 commands.append({'label':label,'argv':[str(x) for x in cmd],'started_utc':utc,'wall_seconds':time.monotonic()-t,'exit_code':r.returncode});write('commands.json',commands)
 if r.returncode:raise RuntimeError(f'{label} failed; retained output at {out}')
case='xlsx_file_selected_cell'
specs=[]
def args(n,c=case,samples=500,warmup=20,alloc=False):
 specs.append({'name':n,'case':c,'samples':samples,'warmup':warmup,'allocator':alloc})
 return [bins['litchi-perf-baseline-alloc' if alloc else 'litchi-perf-baseline'],'--case',c,'--filesystem-cache','warm','--xlsx-cell-crud-shape','medium','--warmup',warmup,'--samples',samples,'--json',out/(n+'.json'),'--corpus-manifest',out/(n+'.catalog.json')]
shutil.copy2(__file__,out/'capture.py')
write('environment.json',{'source_before':before,'binaries':identity,'build_environment':{'CARGO_TARGET_DIR':'/tmp/litchi-goal-0409-target','CARGO_BUILD_JOBS':'4','CARGO_INCREMENTAL':'0','CARGO_PROFILE_RELEASE_DEBUG':'1','RUSTFLAGS':flags,'RUSTUP_TOOLCHAIN':'1.98.1'},'runtime_environment':{'RUSTFLAGS':flags,'RUSTUP_TOOLCHAIN':'1.98.1','DEBUGINFOD_URLS':''},'cpu':2,'performance_claim':'none','clean_abba_claim_eligible':False,'external_profile_scope':'whole process and inherited children, including setup and full semantic verification'})
for n,cmd in [('rustc',['rustc','+1.98.1','-vV']),('cpu',['lscpu']),('memory',['free','-b']),('perf-version',['perf','--version']),('flamegraph-version',['flamegraph','--version']),('normal-build-id',['readelf','-n',bins['litchi-perf-baseline']]),('allocator-build-id',['readelf','-n',bins['litchi-perf-baseline-alloc']])]:run(n,cmd)
run('selected-normal',['taskset','-c','2','/usr/bin/time','-v','-o',out/'selected-normal.time.txt',*args('selected-normal')])
run('selected-allocator',['taskset','-c','2',*args('selected-allocator',samples=30,warmup=3,alloc=True)])
run('selected-perf-stat',['taskset','-c','2','perf','stat','-x,','-o',out/'selected-perf-stat.csv','-e','cycles,instructions,branches,branch-misses,cache-misses,page-faults','--',*args('selected-perf-stat',samples=100,warmup=5)])
run('selected-native-l2',['taskset','-c','2','perf','stat','-x,','-o',out/'selected-native-l2.csv','-e','l2_cache_req_stat.dc_access_in_l2,l2_cache_req_stat.dc_hit_in_l2,l2_cache_req_stat.ls_rd_blk_c','--',*args('selected-native-l2',samples=100,warmup=5)])
run('selected-perf-record',['taskset','-c','2','perf','record','--no-buildid-cache','-e','cycles:u','-F','999','--call-graph','fp,127','-o',out/'perf.data','--',*args('selected-perf-record',samples=300,warmup=10)])
for label,children in [('perf-self','--no-children'),('perf-inclusive','--children')]:run(label,['perf','report','--stdio',children,'--call-graph=graph,0.5,127,caller,function,percent','--percent-limit=0.5','-i',out/'perf.data'])
run('perf-script',['perf','script','-i',out/'perf.data'])
run('flamegraph',['flamegraph','--perfdata',out/'perf.data','--no-inline','--deterministic','--title','XLSX selected cell: whole-process and child CPU samples','--output',out/'cpu-flamegraph.svg'])
for n,c in [('source-edit','xlsx_source_backed_cell_values_one_edit_save'),('eager-edit','xlsx_eager_cell_values_one_edit_save')]:run(n,['taskset','-c','2','/usr/bin/time','-v','-o',out/(n+'.time.txt'),*args(n,c=c)])
sys.path.insert(0,str(repo));from tools import perf_compare,validate_perf_corpus_binding
validation=[]
for spec in specs:
 n=spec['name'];d=json.loads((out/(n+'.json')).read_text());validate_perf_corpus_binding.validate_paths(out/(n+'.json'),out/(n+'.catalog.json'));perf_compare.validate_parallel_metrics(d)
 assert len(d['results'])==1; r=d['results'][0];assert r['case']==spec['case'];assert len(r['elapsed_ns']['samples'])==spec['samples'];assert d['configuration']['warmup_iterations_per_case']==spec['warmup'];assert d['environment']['git_revision']==before['revision'];assert d['environment']['git_worktree_dirty'] is True;assert d['environment']['cpu_affinity']=='2';assert d['environment']['rustflags']==flags
 key='litchi-perf-baseline-alloc' if spec['allocator'] else 'litchi-perf-baseline';assert d['binary_identity']['binary_sha256']==identity[key]['sha256'];assert d['tool']['binary']==key;assert d['tool']['instrumentation']==('system_allocator_operation_scoped' if spec['allocator'] else 'none')
 assert r['corpus']['archive_sha256']=='dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036'
 if spec['case']==case:
  perf_compare._validate_operation_metrics(r['operation_metrics'],n,r['elapsed_ns']['samples'],d['schema_version'],elapsed_sample_order=r['elapsed_ns']['sample_order'])
  assert r['cache_state']=='warm';fs=d['filesystem_evidence'];assert len(fs)==1;fs=fs[0];assert fs['fresh_child_per_sample'] is True;assert fs['sample_count']==spec['samples'];assert len(fs['samples'])==spec['samples']
  pids=set()
  for raw in fs['samples']:
   pids.add(raw['child_process_id']);v=raw['xlsx_selected_cell'];assert v['canonical_sheet_name']=='Bench01' and v['cell_address']=='M29' and v['lexical_value']=='1028012';assert v['digest']=='36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1'
  assert len(pids)==spec['samples']
  if spec['allocator']:assert r['operation_metrics']['allocation']['status']=='measured'
 validation.append({**spec,'status':'passed'})
assert state()==before;assert all(sha(bins[n])==v['sha256'] for n,v in identity.items())
write('validation.json',{'reports':validation,'source_after':state(),'performance_claim':'none','note':'Whole-process attribution, PMU viability, and edit phase evidence require review; no comparison claim is authorized.'})
for f in list(out.glob('*.json'))+[out/'perf-script.stdout']:
 if f.name=='commands.json':continue
 run('compress-'+f.name,['zstd','-q','-f',f,'-o',str(f)+'.zst']);run('test-'+f.name,['zstd','-q','-t',str(f)+'.zst']);assert hashlib.sha256(subprocess.check_output(['zstd','-q','-dc',str(f)+'.zst'])).hexdigest()==sha(f)
write('artifacts.json',[{'path':p.name,'bytes':p.stat().st_size,'sha256':sha(p)} for p in sorted(out.iterdir()) if p.is_file() and p.name!='artifacts.json'])
print('Capture validated:',out)
