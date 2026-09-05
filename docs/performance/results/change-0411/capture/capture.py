#!/usr/bin/env python3
"""Single-revision XLS diagnostic capture; never infer cross-family speedup."""
import argparse,datetime,hashlib,json,os,subprocess,time
from pathlib import Path

p=argparse.ArgumentParser()
p.add_argument('phase',choices=['normal','allocator','profile','counters'])
args=p.parse_args()
out=Path('/tmp/litchi-goal-0411-capture')
out.mkdir(exist_ok=True)
wt=Path('/tmp/litchi-goal-0411-worktree')
bins=Path('/tmp/litchi-goal-0411-binaries')
identity=json.loads((bins/'identity.json').read_text())
env=os.environ.copy()
env.update(RUSTUP_TOOLCHAIN='1.98.1',RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',DEBUGINFOD_URLS='')
cases=['xls_semantic_open','xls_source_backed_open','xls_eager_open_list_worksheets','xls_source_backed_open_list_worksheets','xls_eager_open_one_cell','xls_source_backed_open_one_cell']
def state():
 return dict(revision=subprocess.check_output(['git','rev-parse','HEAD'],cwd=wt,text=True).strip(),status=subprocess.check_output(['git','status','--porcelain=v1'],cwd=wt,text=True).strip())
before=state()
assert before==dict(revision=identity['revision'],status='')
commands_path=out/'commands.json'
commands=json.loads(commands_path.read_text()) if commands_path.exists() else []
def run(label,argv):
 assert label not in [c['label'] for c in commands],label
 argv=list(map(str,argv));stamp=datetime.datetime.now(datetime.timezone.utc).isoformat();start=time.monotonic()
 with (out/(label+'.stdout')).open('wb') as stdout,(out/(label+'.stderr')).open('wb') as stderr:
  r=subprocess.Popen(argv,cwd=wt,env=env,stdout=stdout,stderr=stderr)
  r.wait()
 commands.append(dict(label=label,argv=argv,cwd=str(wt),started_utc=stamp,wall_seconds=time.monotonic()-start,exit_code=r.returncode,launcher_process_id=r.pid))
 commands_path.write_text(json.dumps(commands,indent=2)+'\n')
 print(label,r.returncode,flush=True)
 if r.returncode:raise subprocess.CalledProcessError(r.returncode,argv)
 assert state()==before
def benchmark(label,selected,samples,warmup,allocator=False):
 binary=bins/('litchi-perf-baseline-alloc' if allocator else 'litchi-perf-baseline')
 assert hashlib.sha256(binary.read_bytes()).hexdigest()==identity['binaries'][binary.name]['sha256']
 return [binary,'--filesystem-cache','warm','--case',','.join(selected),'--samples',samples,'--warmup',warmup,'--json',out/(label+'.json'),'--corpus-manifest',out/(label+'.catalog.json')]
if args.phase in ('normal','allocator'):
 allocator=args.phase=='allocator'
 for i in range(1,3 if allocator else 5):
  label=f'{args.phase}-{i}'
  run(label,['taskset','-c','2','/usr/bin/time','-v','-o',out/(label+'.time.txt'),*benchmark(label,cases,30 if allocator else 500,3 if allocator else 20,allocator)])
else:
 for family in ['eager','source_backed']:
  case=f'xls_{family}_open_one_cell'
  if args.phase=='profile':
   label=family+'-profile'
   run(label,['taskset','-c','2','perf','record','--no-buildid-cache','-e','cycles:u','-F','999','--call-graph','fp,127','-o',out/(label+'.data'),'--',*benchmark(label,[case],1000,20)])
   run(label+'-script',['perf','script','--no-inline','-i',out/(label+'.data')])
   run(label+'-self',['perf','report','--stdio','--no-inline','--no-children','--call-graph=none','--percent-limit=0','-i',out/(label+'.data')])
   run(label+'-flamegraph',['flamegraph','--perfdata',out/(label+'.data'),'--no-inline','--deterministic','--title',family+' XLS one cell: whole process','--output',out/(label+'.svg')])
  else:
   for repetition in range(1,4):
    label=f'{family}-stat-{repetition}'
    events='task-clock,cycles,instructions,branches,branch-misses,page-faults,context-switches,cpu-migrations,l2_cache_req_stat.dc_access_in_l2,l2_cache_req_stat.dc_hit_in_l2'
    run(label,['taskset','-c','2','perf','stat','--no-big-num','-x,','-e',events,'-o',out/(label+'.csv'),'--',*benchmark(label,[case],300,10)])
