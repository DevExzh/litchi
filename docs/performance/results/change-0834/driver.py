"""Root-owned phased harness repair; each build/run binds its actual source."""
from pathlib import Path
import hashlib,json,os,shutil,subprocess,sys,time
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
TARGET=ROOT.parent/'litchi-target-0834'
SCRATCH=ROOT.parent/'litchi-fs-0834'
TOOL='tools/perf-baseline/Cargo.toml'
CASES=['opc_file_eager_open','opc_file_source_open','opc_file_eager_one_part_atomic_save',
       'opc_file_source_one_part_atomic_save','pptx_file_eager_open_selected_slide_lifecycle',
       'pptx_file_source_open_selected_slide_lifecycle']
ALLOWED={'tools/perf-baseline/src/filesystem.rs','tools/perf-baseline/src/filesystem/aligned_zip.rs','tools/perf-baseline/README.md'}

def read(p):return json.loads(Path(p).read_text())
def sha(p):
 h=hashlib.sha256()
 with Path(p).open('rb') as f:
  for b in iter(lambda:f.read(1048576),b''):h.update(b)
 return h.hexdigest()
def write(p,v):
 Path(p).parent.mkdir(parents=True,exist_ok=True)
 with Path(p).open('x') as f:json.dump(v,f,indent=2,sort_keys=True);f.write('\n')
def desc(p):
 p=Path(p);assert p.is_file() and not p.is_symlink();return dict(path=str(p),bytes=p.stat().st_size,sha256=sha(p))
def output(a):return subprocess.check_output(a,cwd=ROOT,text=True).strip()
def env():
 e=os.environ.copy();assert not e.get('RUSTFLAGS') and not e.get('CARGO_ENCODED_RUSTFLAGS')
 e.update(CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_DEV_DEBUG='0',RUSTDOCFLAGS='-D warnings',PYTHONDONTWRITEBYTECODE='1',LC_ALL='C',TZ='UTC');return e
def inventory():
 names=set(output(['git','ls-files','crates','tools','Cargo.toml','Cargo.lock','.cargo','clippy.toml','rustfmt.toml','rust-toolchain.toml','docs/adr','docs/GOAL.md','docs/CRUD_Scenario_Checklist.md']).splitlines())
 names.update(n for n in ALLOWED if (ROOT/n).is_file())
 return {n:sha(ROOT/n) for n in sorted(names)}
def invariants():
 o=read(P/'origin.json');assert output(['git','rev-parse','HEAD'])==o['base']
 base=read(P.parent/'change-0833/prepare.json')['source']
 current=inventory()
 assert all(v==base.get(n) for n,v in current.items() if n not in ALLOWED),'unowned source changed'
 assert all(sha(ROOT/n)==v for n,v in o['normative'].items())
 assert all(sha(ROOT/n)==v for n,v in o['unrelated'].items())
 return current

def prepare():
 current=invariants();assert not TARGET.exists() and not SCRATCH.exists();assert 12 in os.sched_getaffinity(0)
 for path in (TARGET,SCRATCH):
  path.mkdir();write(path/'.owner-0834.json',dict(packet=str(P),path=str(path)))
 write(P/'host.json',dict(base=read(P/'origin.json')['base'],rustc=output(['rustc','-Vv']),cargo=output(['cargo','-V']),cpu=Path('/proc/cpuinfo').read_text(),memory=Path('/proc/meminfo').read_text(),affinity=sorted(os.sched_getaffinity(0)),filesystem=output(['findmnt','-T',str(ROOT),'-n','-o','FSTYPE,OPTIONS']),fincore=desc(Path(shutil.which('fincore')).resolve()),environment={k:env().get(k) for k in ['PATH','CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_DEV_DEBUG','RUSTDOCFLAGS','RUSTFLAGS','LD_PRELOAD','LD_LIBRARY_PATH']},started_unix=time.time()))
 print('prepare PASS',len(current),flush=True)
def freeze(stage):
 current=invariants();assert stage in ['diagnostic','repaired']
 for n in ALLOWED:
  if (ROOT/n).is_file():
   dest=P/'sources'/stage/n;dest.parent.mkdir(parents=True,exist_ok=True)
   with dest.open('xb') as f:f.write((ROOT/n).read_bytes())
 write(P/f'freeze-{stage}.json',dict(source=current,driver=desc(P/'driver.py'),origin=desc(P/'origin.json'),stage=stage,created_unix=time.time()))
 print('freeze',stage,'PASS',flush=True)
def check(stage):
 assert invariants()==read(P/f'freeze-{stage}.json')['source'],'stage source changed'
 assert desc(P/'driver.py')==read(P/f'freeze-{stage}.json')['driver']
 for path in (TARGET,SCRATCH):assert read(path/'.owner-0834.json')==dict(packet=str(P),path=str(path))
def run(stage,label,argv):
 check(stage);out=P/'commands'/label;out.mkdir(parents=True,exist_ok=False)
 start=time.time();write(out/'started.json',dict(argv=argv,cwd=str(ROOT),started_unix=start,freeze=desc(P/f'freeze-{stage}.json')))
 code=None;error=None
 with (out/'output.log').open('xb') as f:
  try:code=subprocess.run(argv,cwd=ROOT,env=env(),stdout=f,stderr=subprocess.STDOUT).returncode
  except Exception as e:error=repr(e)
 row=dict(argv=argv,exit_code=code,error=error,started_unix=start,finished_unix=time.time(),log=desc(out/'output.log'),freeze=desc(P/f'freeze-{stage}.json'))
 write(out/'receipt.json',row);print(label,code,error,flush=True);return row

def build(stage):
 r=run(stage,f'build-{stage}',['cargo','build','--offline','--locked','--release','--manifest-path',TOOL,'--bin','litchi-perf-baseline']);assert r['exit_code']==0
 dest=TARGET/'retained'/stage/'litchi-perf-baseline';dest.parent.mkdir(parents=True,exist_ok=False);shutil.copy2(TARGET/'release/litchi-perf-baseline',dest)
 write(P/f'build-{stage}.json',dict(binary=desc(dest),receipt=desc(P/f'commands/build-{stage}/receipt.json')))
def workload(stage,label,case,state,samples=1,warmup=0):
 binary=read(P/f'build-{stage}.json')['binary'];assert desc(binary['path'])==binary
 report=P/f'{label}.json';assert not report.exists()
 argv=['taskset','-c','12',binary['path'],'--case',case,'--samples',str(samples),'--warmup',str(warmup),'--filesystem-cache',state,'--filesystem-root',str(SCRATCH),'--json',str(report)]
 r=run(stage,label,argv)
 return dict(case=case,state=state,exit_code=r['exit_code'],report=desc(report) if report.exists() else None,receipt=desc(P/f'commands/{label}/receipt.json'))
def diagnose():
 rows=[]
 for state in ['warm','cold-verified']:
  rows.append(workload('diagnostic',f'diagnostic-pptx-{state}',CASES[5],state))
 write(P/'diagnostic-results.json',dict(rows=rows,performance_claim='none'))
def quality():
 prefix=['--offline','--locked','--manifest-path',TOOL]
 gates=[('fmt',['cargo','fmt','--manifest-path',TOOL,'--','--check']),
 ('check',['cargo','check',*prefix,'--all-features','--all-targets']),
 ('test',['cargo','test',*prefix,'--lib','--features','allocator-metrics,ordinary-save-process-metrics','--','--test-threads=2']),
 ('clippy',['cargo','clippy',*prefix,'--all-features','--all-targets','--','-D','warnings']),
 ('doc',['cargo','doc',*prefix,'--all-features','--no-deps']),
 ('boundaries',['python3','-B','tools/check_crate_boundaries.py'])]
 for name,argv in gates:assert run('repaired','quality-'+name,argv)['exit_code']==0,name
 write(P/'quality.json',dict(status='pass',gates=[n for n,_ in gates],reused=False))
def qualify():
 assert read(P/'quality.json')['status']=='pass'
 rows=[workload('repaired',f'qualification-{i:02}',c,'warm,cold-verified') for i,c in enumerate(CASES)]
 rows.append(workload('repaired','qualification-opc-pair',','.join(CASES[2:4]),'warm,cold-verified'))
 write(P/'qualification.json',dict(rows=rows,status='commands_pass' if all(x['exit_code']==0 for x in rows) else 'failed'))
 assert all(x['exit_code']==0 for x in rows)
if __name__=='__main__':
 cmd=sys.argv[1];args=sys.argv[2:];assert cmd in ['prepare','freeze','build','diagnose','quality','qualify'];globals()[cmd](*args)
