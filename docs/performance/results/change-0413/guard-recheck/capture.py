#!/usr/bin/env python3
"""0413 sequential production CFB scratch comparison."""
import argparse,datetime,hashlib,json,os,subprocess,time
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('phase',choices=['normal','allocator','guard-normal','guard-allocator','profile','counters']);args=p.parse_args()
out=Path('/tmp/litchi-goal-0413-guard-recheck');out.mkdir(exist_ok=True)
env=os.environ.copy();env.update(RUSTUP_TOOLCHAIN='1.98.1',RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',DEBUGINFOD_URLS='')
cases=json.loads((out/'protocol.json').read_text())['cases']
owned=['xls_owned_source_open','xls_owned_source_open_list_worksheets','xls_owned_source_open_one_cell']
commands_path=out/'commands.json';commands=json.loads(commands_path.read_text()) if commands_path.exists() else []
def run(label,variant,selected,samples,warmup,allocator=False,profile=False,counters=False,guard=False):
 assert label not in [c['label'] for c in commands]
 wt=Path('/tmp/litchi-goal-0413-'+variant+'-worktree');bins=Path('/tmp/litchi-goal-0413-'+variant+'-binaries')
 identity=json.loads((bins/'identity.json').read_text())
 assert identity['revision']==json.loads((out/'protocol.json').read_text())[variant+'_revision']
 def state():return dict(revision=subprocess.check_output(['git','rev-parse','HEAD'],cwd=wt,text=True).strip(),status=subprocess.check_output(['git','status','--porcelain=v1'],cwd=wt,text=True).strip())
 expected=dict(revision=identity['revision'],status='');assert state()==expected
 binary=bins/('litchi-perf-baseline-alloc' if allocator else 'litchi-perf-baseline')
 assert hashlib.sha256(binary.read_bytes()).hexdigest()==identity['binaries'][binary.name]['sha256']
 argv=[str(binary),'--filesystem-cache','warm','--case',','.join(selected),'--samples',str(samples),'--warmup',str(warmup),'--json',str(out/(label+'.json')),'--corpus-manifest',str(out/(label+'.catalog.json'))]
 if guard:argv.extend(['--shape','tiny,few-large','--payload','incompressible'])
 if profile:argv=['perf','record','--no-buildid-cache','-e','cycles:u','-F','999','--call-graph','fp,127','-o',str(out/(label+'.data')),'--',*argv]
 elif counters:argv=['perf','stat','--no-big-num','-x,','-e',json.loads((out/'protocol.json').read_text())['counters']['events'],'-o',str(out/(label+'.csv')),'--',*argv]
 else:argv=['/usr/bin/time','-v','-o',str(out/(label+'.time.txt')),*argv]
 argv=['taskset','-c','2',*argv]
 stamp=datetime.datetime.now(datetime.timezone.utc).isoformat();start=time.monotonic()
 with (out/(label+'.stdout')).open('wb') as stdout,(out/(label+'.stderr')).open('wb') as stderr:
  r=subprocess.Popen(argv,cwd=wt,env=env,stdout=stdout,stderr=stderr);r.wait()
 commands.append(dict(label=label,variant=variant,revision=identity['revision'],source_status='',protocol_sha256=hashlib.sha256((out/'protocol.json').read_bytes()).hexdigest(),binary_sha256=identity['binaries'][binary.name]['sha256'],argv=argv,cwd=str(wt),started_utc=stamp,wall_seconds=time.monotonic()-start,exit_code=r.returncode,launcher_process_id=r.pid))
 commands_path.write_text(json.dumps(commands,indent=2)+'\n');print(label,r.returncode,flush=True)
 if r.returncode:raise subprocess.CalledProcessError(r.returncode,argv)
 assert state()==expected
if args.phase in ('normal','allocator','guard-normal','guard-allocator'):
 alloc='allocator' in args.phase;guard=args.phase.startswith('guard-')
 for i,variant in enumerate(['control','candidate','candidate','control'],1):run(f'{args.phase}-{i}',variant,['cfb_open'] if guard else cases,30 if alloc else 1000,3 if alloc else 50,alloc,guard=guard)
elif args.phase=='counters':
 for i,variant in enumerate(['control','candidate','candidate','control'],1):run(f'stat-{i}',variant,[owned[-1]],3000,10,counters=True)
else:
 for variant in ['control','candidate']:run(variant+'-profile',variant,[owned[-1]],10000,20,profile=True)
