"""Serial paired release measurement; preserve every build/run and source hash."""
from pathlib import Path
import hashlib,json,os,shutil,subprocess,time,platform
P=Path(__file__).resolve().parent
AFTER=P.parents[3]
BEFORE=Path('/home/zhuhe/code/litchi')
REFS={'before':'25c3b27ba4','after':'836120efe7'}
ROOTS={'before':BEFORE,'after':AFTER}
TARGETS={leg:Path('/home/zhuhe/code/litchi-target-0774-release-'+leg) for leg in ROOTS}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def source(root,ref):
 subprocess.run(['git','diff','--quiet',ref,'--','crates','tools/perf-baseline','Cargo.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=root,check=True)
 names=subprocess.check_output(['git','ls-files','-z','crates','tools/perf-baseline','Cargo.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=root).decode().split('\0')
 return {n:sha(root/n) for n in names if n}|{'Cargo.lock':sha(root/'Cargo.lock')}
if __name__=='__main__':
 out=P/'measure-1';assert not out.exists();out.mkdir();inputs={leg:source(root,REFS[leg]) for leg,root in ROOTS.items()}
 for leg,data in inputs.items():write(out/f'source-{leg}.json',data)
 assert inputs['before']['Cargo.lock']==inputs['after']['Cargo.lock']
 assert 12 in os.sched_getaffinity(0)
 env=os.environ|{'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','RUSTUP_TOOLCHAIN':'1.95.0','PYTHONDONTWRITEBYTECODE':'1'}
 write(out/'environment.json',{'platform':platform.platform(),'cpu':12,'allowed_cpus':sorted(os.sched_getaffinity(0)),'rustc':subprocess.check_output(['rustc','-Vv']).decode(),'profile':'release opt-level3, LTO=true, panic=abort, debug=false','environment':{k:env.get(k) for k in ['CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTUP_TOOLCHAIN','RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS']},'references':REFS})
 rows=[];binaries={}
 for leg,root in ROOTS.items():
  probe=out/leg;probe.mkdir();(probe/'src').mkdir();shutil.copy2(P/'probe-src/main.rs',probe/'src/main.rs');(probe/'Cargo.toml').write_text((P/'probe-src/Cargo.toml.template').read_text().replace('@SRC@',str(root)))
  buildenv=env|{'CARGO_TARGET_DIR':str(TARGETS[leg])}
  if leg=='before':
   cmd=['cargo','generate-lockfile','--offline','--manifest-path',str(probe/'Cargo.toml')]
   with (out/'lock.log').open('w') as f:subprocess.run(cmd,cwd=AFTER,env=buildenv,stdout=f,stderr=subprocess.STDOUT,check=True)
  else:shutil.copy2(out/'before/Cargo.lock',probe/'Cargo.lock')
  cmd=['cargo','build','--release','--offline','--locked','--manifest-path',str(probe/'Cargo.toml')];log=out/f'build-{leg}.log';start=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=AFTER,env=buildenv,stdout=f,stderr=subprocess.STDOUT)
  rows.append({'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'log':log.name,'sha256':sha(log)});write(out/'build.json',rows);assert r.returncode==0
  binary=TARGETS[leg]/'release/xls-formula-probe';binaries[leg]={'path':str(binary),'bytes':binary.stat().st_size,'sha256':sha(binary),'lock_sha256':sha(probe/'Cargo.lock')}
 write(out/'binaries.json',binaries)
 for leg in ROOTS:assert source(ROOTS[leg],REFS[leg])==inputs[leg]
 runs=[]
 # Four process pairs, alternating order, with numeric control in every process.
 for case in ['formula','numeric']:
  for index,leg in enumerate(['before','after','after','before','before','after','after','before']):
   prefix=f'{case}-{index}-{leg}';stdout=out/(prefix+'.json');stderr=out/(prefix+'.stderr');rss=out/(prefix+'.rss');binary=Path(binaries[leg]['path']);assert sha(binary)==binaries[leg]['sha256']
   cmd=['/usr/bin/time','-f','%M','-o',str(rss),'taskset','-c','12',str(binary),'--case',case,'--samples','9','--warmup','2'];start=time.time()
   with stdout.open('w') as o,stderr.open('w') as e:r=subprocess.run(cmd,cwd=AFTER,env=env,stdout=o,stderr=e)
   runs.append({'leg':leg,'case':case,'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'stdout':stdout.name,'stdout_sha256':sha(stdout),'stderr':stderr.name,'stderr_sha256':sha(stderr),'rss':rss.name,'rss_sha256':sha(rss)});write(out/'runs.json',runs);assert r.returncode==0
   print(prefix+' done',flush=True)
 for leg in ROOTS:assert source(ROOTS[leg],REFS[leg])==inputs[leg]
 write(out/'complete.json',{'source_unchanged':True,'serial':True,'runs':len(runs)})
