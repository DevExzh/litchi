#!/usr/bin/env python3
"""Capture one committed, explicitly instrumented OPC materialization baseline."""
import argparse, datetime, hashlib, json, os, shutil, subprocess, sys
from pathlib import Path
if not __debug__: raise RuntimeError('Run without python -O; validation assertions are required')
p=argparse.ArgumentParser();p.add_argument('--bin-dir',type=Path,required=True);p.add_argument('--revision',required=True);a=p.parse_args()
repo=Path.cwd();out=Path('/tmp/litchi-goal-0408');assert not out.exists();out.mkdir()
def digest(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def write(name,data): (out/name).write_text(json.dumps(data,indent=2)+'\n')
def git(*args):return subprocess.check_output(['git',*args],text=True).strip()
def state():return {'revision':git('rev-parse','HEAD'),'status':git('status','--porcelain=v1'),'goal_sha256':digest(repo/'docs/GOAL.md')}
before=state();assert before['revision']==a.revision;assert before['status']=='?? docs/GOAL.md',before
flags='-C force-frame-pointers=yes -C force-unwind-tables=yes'
env=os.environ.copy();env.update(RUSTUP_TOOLCHAIN='1.98.1',RUSTFLAGS=flags,DEBUGINFOD_URLS='')
bins={n:(a.bin_dir/n).resolve() for n in ['litchi-perf-baseline','litchi-perf-baseline-alloc']}
identities={n:{'path':str(f),'sha256':digest(f),'bytes':f.stat().st_size} for n,f in bins.items()}
commands=[]
def run(label,cmd,check=True):
 import time
 start=datetime.datetime.now(datetime.timezone.utc).isoformat();t=time.monotonic()
 with (out/(label+'.stdout')).open('wb') as stdout,(out/(label+'.stderr')).open('wb') as stderr:
  result=subprocess.run([str(v) for v in cmd],env=env,stdout=stdout,stderr=stderr)
 commands.append({'label':label,'argv':[str(v) for v in cmd],'started_utc':start,'wall_seconds':time.monotonic()-t,'exit_code':result.returncode})
 write('commands.json',commands)
 if result.returncode and check:raise RuntimeError(f'{label} exited {result.returncode}; artifacts retained at {out}')
 return result
write('tool-paths.json',{n:shutil.which(n) for n in ['flamegraph','zstd','taskset','readelf','perf','cargo','rustc']})
for n in ['flamegraph','zstd','taskset','readelf']:run(n+'-version',[n,'--version'])
normal='opc_source_materialize';accounted='opc_source_materialize_accounted'
def args(name,case=normal,allocator=False):return [bins['litchi-perf-baseline-alloc' if allocator else 'litchi-perf-baseline'],'--case',case,'--shape','few-large','--payload','incompressible','--warmup','20','--samples','500','--json',out/(name+'.json'),'--corpus-manifest',out/(name+'.catalog.json')]
shutil.copy2(__file__,out/'capture.py')
write('environment.json',{'source_before':before,'binaries':identities,'build_environment':{'RUSTUP_TOOLCHAIN':'1.98.1','RUSTFLAGS':flags,'CARGO_PROFILE_RELEASE_DEBUG':'1','CARGO_TARGET_DIR':'/tmp/litchi-goal-fp-target','CARGO_BUILD_JOBS':'4','CARGO_INCREMENTAL':'0','strip':'Cargo default (none)','split_debuginfo':'Cargo target default'},'runtime_environment':{'RUSTUP_TOOLCHAIN':'1.98.1','RUSTFLAGS':flags,'DEBUGINFOD_URLS':''},'cpu':2,'sample_scope':'in-process operation; external profilers include setup and full verification','performance_claim':'none','clean_abba_claim_eligible':False})
for label,cmd in [('rustc',['rustc','+1.98.1','-vV']),('cargo',['cargo','+1.98.1','-V']),('cpu',['lscpu']),('memory',['free','-b']),('perf-version',['perf','--version']),('normal-sections',['readelf','-S','--wide',bins['litchi-perf-baseline']]),('normal-build-id',['readelf','-n',bins['litchi-perf-baseline']]),('allocator-build-id',['readelf','-n',bins['litchi-perf-baseline-alloc']])]:run(label,cmd)
run('normal',['taskset','-c','2','/usr/bin/time','-v','-o',out/'time-v.txt',*args('normal')])
run('allocator',['taskset','-c','2',*args('allocator',allocator=True)])
run('accounted',['taskset','-c','2',*args('accounted',case=accounted)])
run('perf-stat',['taskset','-c','2','perf','stat','-x,','-o',out/'perf-stat.csv','-e','cycles,instructions,branches,branch-misses,cache-misses,page-faults','--',*args('perf-stat')])
cache_result=run('perf-cache',['taskset','-c','2','perf','stat','-x,','-o',out/'perf-cache.csv','-e','L1-dcache-loads,L1-dcache-load-misses,LLC-loads,LLC-load-misses','--',*args('perf-cache')],check=False)
run('perf-record',['taskset','-c','2','perf','record','--no-buildid-cache','-e','cycles:u','-F','999','--call-graph','fp,127','-o',out/'perf.data','--',*args('perf-record')])
for label,children in [('perf-self','--no-children'),('perf-inclusive','--children')]:run(label,['perf','report','--stdio',children,'--call-graph=graph,0.5,127,caller,function,percent','--percent-limit=0.5','-i',out/'perf.data'])
run('perf-script',['perf','script','-i',out/'perf.data'])
run('flamegraph',['flamegraph','--perfdata',out/'perf.data','--no-inline','--deterministic','--title','OPC materialization: whole-process CPU samples','--output',out/'cpu-flamegraph.svg'])
sys.path.insert(0,str(repo));from tools import perf_compare, validate_perf_corpus_binding
reports=[]
for name,expected_case,alloc in [('normal',normal,False),('allocator',normal,True),('accounted',accounted,False),('perf-stat',normal,False),('perf-cache',normal,False),('perf-record',normal,False)]:
 if name=='perf-cache' and cache_result.returncode != 0:continue
 d=json.loads((out/(name+'.json')).read_text());perf_compare.validate_parallel_metrics(d);assert len(d['results'])==1
 validate_perf_corpus_binding.validate_paths(out/(name+'.json'),out/(name+'.catalog.json'))
 r=d['results'][0];assert r['case']==expected_case;assert len(r['elapsed_ns']['samples'])==500
 assert d['environment']['git_revision']==a.revision;assert d['environment']['git_worktree_dirty'] is True;assert '1.98.1' in d['environment']['rustc_version'];assert d['environment']['cpu_affinity']=='2'
 perf_compare._validate_opc_source_materialize_oracle(d['configuration'],name)
 assert d['configuration']['opc_source_materialize_oracle']=='prepared-part-digest-v1'
 assert d['environment']['rustflags']==flags
 assert d['environment']['allocator']==('CountingSystemAllocator(std::alloc::System)' if alloc else 'Rust system allocator')
 assert r['corpus']['shape']=='few-large' and r['corpus']['payload_kind']=='incompressible'
 assert r['corpus']['archive_sha256']=='a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6'
 key='litchi-perf-baseline-alloc' if alloc else 'litchi-perf-baseline'
 assert d['tool']['binary']==key and d['tool']['profile']=='release'
 assert d['tool']['instrumentation']==('system_allocator_operation_scoped' if alloc else 'none')
 assert d['binary_identity']['profile']=='release' and d['binary_identity']['binary_bytes']==identities[key]['bytes']
 assert d['binary_identity']['binary_sha256']==identities[key]['sha256']
 perf_compare._validate_operation_metrics(r['operation_metrics'],name,r['elapsed_ns']['samples'],d['schema_version'],elapsed_sample_order=r['elapsed_ns']['sample_order'])
 perf_compare._validate_operation_metric_case_binding(r['operation_metrics'],name,r['case'])
 assert r['operation_metrics']['allocation']['status']==('measured' if alloc else 'unavailable')
 assert ('opc_zip' in r['operation_metrics'])==(name=='accounted')
 reports.append({'name':name,'corpus':r['corpus'],'validation':'passed'})
assert all(x['corpus']==reports[0]['corpus'] for x in reports)
after=state();assert after==before
assert all(digest(bins[n])==v['sha256'] for n,v in identities.items())
for f in [x for x in out.glob('*.json') if x.name!='commands.json']+[out/'perf-script.stdout']:
 run('compress-'+f.name,['zstd','-q','-f',f,'-o',str(f)+'.zst'])
 run('test-compression-'+f.name,['zstd','-q','-t',str(f)+'.zst'])
 assert hashlib.sha256(subprocess.check_output(['zstd','-q','-dc',str(f)+'.zst'])).hexdigest()==digest(f)
write('validation.json',{'source_after':after,'reports':reports,'performance_claim':'none','required_commands_succeeded':all(x['exit_code']==0 for x in commands if x['label']!='perf-cache'),'optional_counter_status':[x for x in commands if x['label']=='perf-cache'],'note':'Counter support and caller-unwind quality still require inspection; no performance comparison is authorized.'})
write('artifacts.json',[{'path':str(f.relative_to(out)),'bytes':f.stat().st_size,'sha256':digest(f)} for f in sorted(out.rglob('*')) if f.is_file() and f.name!='artifacts.json'])
print('Captured and validated',out)
