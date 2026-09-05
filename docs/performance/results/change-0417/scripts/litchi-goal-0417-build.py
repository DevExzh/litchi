from pathlib import Path
import subprocess,os,time,json,hashlib,datetime
repo=Path('/home/zhuhe/code/litchi');out=repo/'docs/performance/results/change-0417';tree=Path('/tmp/litchi-goal-0417-worktree');target=repo/'tools/perf-baseline/target'
env=os.environ|{'RUSTUP_TOOLCHAIN':'1.98.1','CARGO_TARGET_DIR':str(target),'CARGO_BUILD_JOBS':'4','CARGO_INCREMENTAL':'0','CARGO_PROFILE_RELEASE_DEBUG':'1','RUSTFLAGS':'-C force-frame-pointers=yes -C force-unwind-tables=yes'}
argv=['cargo','+1.98.1','build','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--release','--features','allocator-metrics','--bin','litchi-perf-baseline','--bin','litchi-perf-baseline-alloc']
start=datetime.datetime.now(datetime.timezone.utc).isoformat();before=time.monotonic()
with (out/'checks/build.log').open('w') as log:p=subprocess.run(argv,cwd=tree,env=env,stdout=log,stderr=subprocess.STDOUT)
d={'revision':subprocess.check_output(['git','rev-parse','HEAD'],cwd=tree,text=True).strip(),'source_status':subprocess.check_output(['git','status','--porcelain'],cwd=tree,text=True),'argv':argv,'cwd':str(tree),'environment':{k:env[k] for k in ['RUSTUP_TOOLCHAIN','CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_DEBUG','RUSTFLAGS']},'exit_code':p.returncode,'started_utc':start,'elapsed_seconds':time.monotonic()-before}
if p.returncode==0:
 import shutil
 d['binaries']={}
 for phase,name in [('normal','litchi-perf-baseline'),('allocator','litchi-perf-baseline-alloc')]:
  dst=Path('/tmp/litchi-goal-0417-'+phase);shutil.copy2(target/'release'/name,dst);d['binaries'][phase]={'path':str(dst),'sha256':hashlib.sha256(dst.read_bytes()).hexdigest(),'bytes':dst.stat().st_size}
(out/'build-identity.json').write_text(json.dumps(d,indent=2)+'\n');print(json.dumps(d));raise SystemExit(p.returncode)
