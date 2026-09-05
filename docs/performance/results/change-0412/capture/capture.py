#!/usr/bin/env python3
"""0412 sequential observer comparison and plain source baseline capture."""
import argparse,datetime,hashlib,json,os,subprocess,time
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('phase',choices=['normal','allocator','owned','owned-allocator','profile','counters']);args=p.parse_args()
out=Path('/tmp/litchi-goal-0412-capture');out.mkdir(exist_ok=True)
env=os.environ.copy();env.update(RUSTUP_TOOLCHAIN='1.98.1',RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',DEBUGINFOD_URLS='')
cases=json.loads((out/'protocol.json').read_text())['cases']
owned=['xls_owned_source_open','xls_owned_source_open_list_worksheets','xls_owned_source_open_one_cell']
commands_path=out/'commands.json';commands=json.loads(commands_path.read_text()) if commands_path.exists() else []
def run(label,variant,selected,samples,warmup,allocator=False,profile=False,counters=False):
 assert label not in [c['label'] for c in commands]
 wt=Path('/tmp/litchi-goal-0412-'+variant+'-worktree');bins=Path('/tmp/litchi-goal-0412-'+variant+'-binaries')
 identity=json.loads((bins/'identity.json').read_text())
 def state():return dict(revision=subprocess.check_output(['git','rev-parse','HEAD'],cwd=wt,text=True).strip(),status=subprocess.check_output(['git','status','--porcelain=v1'],cwd=wt,text=True).strip())
 expected=dict(revision=identity['revision'],status='');assert state()==expected
 binary=bins/('litchi-perf-baseline-alloc' if allocator else 'litchi-perf-baseline')
 assert hashlib.sha256(binary.read_bytes()).hexdigest()==identity['binaries'][binary.name]['sha256']
 argv=[str(binary),'--filesystem-cache','warm','--case',','.join(selected),'--samples',str(samples),'--warmup',str(warmup),'--json',str(out/(label+'.json')),'--corpus-manifest',str(out/(label+'.catalog.json'))]
 if profile:argv=['perf','record','--no-buildid-cache','-e','cycles:u','-F','999','--call-graph','fp,127','-o',str(out/(label+'.data')),'--',*argv]
 elif counters:argv=['perf','stat','--no-big-num','-x,','-e',json.loads((out/'protocol.json').read_text())['counters']['events'],'-o',str(out/(label+'.csv')),'--',*argv]
 else:argv=['/usr/bin/time','-v','-o',str(out/(label+'.time.txt')),*argv]
 argv=['taskset','-c','2',*argv]
 stamp=datetime.datetime.now(datetime.timezone.utc).isoformat();start=time.monotonic()
 with (out/(label+'.stdout')).open('wb') as stdout,(out/(label+'.stderr')).open('wb') as stderr:
  r=subprocess.Popen(argv,cwd=wt,env=env,stdout=stdout,stderr=stderr);r.wait()
 commands.append(dict(label=label,variant=variant,revision=identity['revision'],source_status='',binary_sha256=identity['binaries'][binary.name]['sha256'],argv=argv,cwd=str(wt),started_utc=stamp,wall_seconds=time.monotonic()-start,exit_code=r.returncode,launcher_process_id=r.pid))
 commands_path.write_text(json.dumps(commands,indent=2)+'\n');print(label,r.returncode,flush=True)
 if r.returncode:raise subprocess.CalledProcessError(r.returncode,argv)
 assert state()==expected
if args.phase in ('normal','allocator'):
 alloc=args.phase=='allocator'
 for i,variant in enumerate(['control','candidate','candidate','control'],1):run(f'{args.phase}-{i}',variant,cases,30 if alloc else 500,3 if alloc else 20,alloc)
elif args.phase in ('owned','owned-allocator'):
 alloc=args.phase=='owned-allocator'
 for i in range(1,3 if alloc else 5):run(f'{args.phase}-{i}','candidate',owned,30 if alloc else 500,3 if alloc else 20,alloc)
elif args.phase=='counters':
 for i in range(1,4):run(f'owned-stat-{i}','candidate',[owned[-1]],3000,10,counters=True)
else:run('owned-profile','candidate',[owned[-1]],10000,20,profile=True)
