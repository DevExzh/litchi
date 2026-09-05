import datetime,hashlib,json,os,shutil,subprocess,time
from pathlib import Path
wt=Path('/tmp/litchi-goal-0413-candidate-worktree');out=Path('/tmp/litchi-goal-0413-candidate-checks');bins=Path('/tmp/litchi-goal-0413-candidate-binaries')
bins.mkdir(exist_ok=True)
env=os.environ.copy();build_env=dict(RUSTUP_TOOLCHAIN='1.98.1',CARGO_TARGET_DIR='/tmp/litchi-goal-0413-target',CARGO_BUILD_JOBS='4',CARGO_INCREMENTAL='0',CARGO_PROFILE_RELEASE_DEBUG='1',RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes');env.update(build_env)
def git(*args):return subprocess.check_output(['git',*args],cwd=wt,text=True).strip()
assert git('status','--porcelain=v1')==''
argv=['cargo','+1.98.1','build','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--release','--features','allocator-metrics','--bin','litchi-perf-baseline','--bin','litchi-perf-baseline-alloc']
stamp=datetime.datetime.now(datetime.timezone.utc).isoformat();start=time.monotonic()
with (out/'release-build.log').open('wb') as log:r=subprocess.run(argv,cwd=wt,env=env,stdout=log,stderr=subprocess.STDOUT)
identity=dict(revision=git('rev-parse','HEAD'),source_status=git('status','--porcelain=v1'),build_environment=build_env,build_command=argv,build_cwd=str(wt),started_utc=stamp,wall_seconds=time.monotonic()-start,exit_code=r.returncode)
(out/'release-build.json').write_text(json.dumps(identity,indent=2)+'\n');r.check_returncode();assert identity['source_status']==''
identity['binaries']={}
for name in ['litchi-perf-baseline','litchi-perf-baseline-alloc']:
 p=bins/name;shutil.copy2(Path(build_env['CARGO_TARGET_DIR'])/'release'/name,p)
 identity['binaries'][name]=dict(bytes=p.stat().st_size,sha256=hashlib.sha256(p.read_bytes()).hexdigest())
for name,argv in [('rustc',['rustc','+1.98.1','-vV']),('cargo',['cargo','+1.98.1','-V']),('kernel',['uname','-a']),('cpu',['lscpu']),('perf',['perf','--version'])]:
 identity[name]=subprocess.check_output(argv,env=env,text=True)
(bins/'identity.json').write_text(json.dumps(identity,indent=2)+'\n');print('Built frozen normal/allocator binaries',identity['revision'],flush=True)
