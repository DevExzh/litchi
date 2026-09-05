from pathlib import Path
import subprocess,json,datetime,sys,hashlib
repo=Path('/home/zhuhe/code/litchi');root=repo/'docs/performance/results/change-0417';tree=Path('/tmp/litchi-goal-0417-worktree')
matrix=json.loads((root/'matrix.json').read_text());build=json.loads((root/'build-identity.json').read_text());phase=sys.argv[1]
assert phase in ('preflight','normal','allocator')
assert subprocess.check_output(['git','status','--porcelain'],cwd=tree,text=True)==''
assert build['exit_code']==0
binary=build['binaries']['allocator' if phase=='allocator' else 'normal'];assert hashlib.sha256(Path(binary['path']).read_bytes()).hexdigest()==binary['sha256']
manifest=root/('checks/preflight-capture.json' if phase=='preflight' else 'capture.json')
capture=json.loads(manifest.read_text()) if manifest.exists() else {'revision':build['revision'],'binaries':build['binaries'],'runs':[]}
assert not any(r['phase']==phase for r in capture['runs']), 'refuse overwrite/restart'
for repeat in (range(1,2) if phase=='preflight' else range(1,3)):
 jobs=matrix['jobs'] if repeat==1 else list(reversed(matrix['jobs']))
 for job in jobs:
  selector=job['selector'];samples,warmups=(1,0) if phase=='preflight' else ((500,20) if phase=='normal' else (30,3))
  folder=root/phase;folder.mkdir(exist_ok=True);stem=f'{repeat}-{selector}';report=folder/(stem+'.json');catalog=folder/(stem+'.catalog.json');timing=folder/(stem+'.time.txt')
  args=[binary['path'],'--case',selector,*matrix['common_flags'],'--samples',str(samples),'--warmup',str(warmups),'--json',str(report),'--corpus-manifest',str(catalog)]
  argv=['taskset','-c','2','/usr/bin/time','-v','-o',str(timing),*args]
  started=datetime.datetime.now(datetime.timezone.utc).isoformat()
  p=subprocess.run(argv,cwd=tree,capture_output=True,text=True)
  rec={'phase':phase,'repeat':repeat,'selector':selector,'report':report.relative_to(root).as_posix(),'catalog':catalog.relative_to(root).as_posix(),'argv':argv,'exit_code':p.returncode,'started_utc':started,'finished_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'stderr':p.stderr,'stdout':p.stdout,'time_v':timing.relative_to(root).as_posix()}
  capture['runs'].append(rec);manifest.write_text(json.dumps(capture,indent=2)+'\n');print(phase,repeat,selector,p.returncode,flush=True)
  if p.returncode and phase!='preflight':raise SystemExit(p.returncode)
