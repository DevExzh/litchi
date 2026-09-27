"""Serial paired release measurement; preserve every build/run and source hash."""
from pathlib import Path
import hashlib,json,os,shutil,subprocess,time,platform
P=Path(__file__).resolve().parent
AFTER=P.parents[3]
BEFORE=Path('/home/zhuhe/code/litchi')
REFS={'before':'92dfbe0bd5','after':'fbb579914c'}
ROOTS={'before':BEFORE,'after':AFTER}
TARGETS={leg:Path('/home/zhuhe/code/litchi-target-0775-release-'+leg) for leg in ROOTS}
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def source(root,ref):
 subprocess.run(['git','diff','--quiet',ref,'--','crates','tools/perf-baseline','Cargo.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=root,check=True)
 names=subprocess.check_output(['git','ls-files','-z','crates','tools/perf-baseline','Cargo.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=root).decode().split('\0')
 return {n:sha(root/n) for n in names if n}|{'Cargo.lock':sha(root/'Cargo.lock')}
if __name__=='__main__':
 out=P/'measure-2';assert not out.exists();out.mkdir();probe_inputs={n:sha(P/'probe-src'/n) for n in ['main.rs','Cargo.toml.template']};write(out/'probe-inputs.json',probe_inputs);inputs={leg:source(root,REFS[leg]) for leg,root in ROOTS.items()}
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
  binary=TARGETS[leg]/'release/mce-stream-probe';binaries[leg]={'path':str(binary),'bytes':binary.stat().st_size,'sha256':sha(binary),'lock_sha256':sha(probe/'Cargo.lock')}
 write(out/'binaries.json',binaries)
 assert {n:sha(P/'probe-src'/n) for n in probe_inputs}==probe_inputs
 assert sha(out/'before/src/main.rs')==sha(out/'after/src/main.rs')==probe_inputs['main.rs']
 for leg in ROOTS:assert source(ROOTS[leg],REFS[leg])==inputs[leg]
 runs=[]
 fixture_roots=['test-data/ooxml/docx','test-data/ooxml/xlsx','test-data/ooxml/pptx']
 def fixtures():return {str(f.relative_to(AFTER)):sha(f) for root in fixture_roots for f in sorted((AFTER/root).rglob('*')) if f.is_file()}
 fixture_files=fixtures();write(out/'fixtures.json',fixture_files)
 for leg in ['before','after']:
  binary=Path(binaries[leg]['path']);assert sha(binary)==binaries[leg]['sha256']
  report=out/f'differential-{leg}.json';log=out/f'differential-{leg}.log'
  cmd=['taskset','-c','12',str(binary),'differential','--generated','4096','--aliasing','4096','--json',str(report),*fixture_roots];start=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=AFTER,env=env,stdout=f,stderr=subprocess.STDOUT)
  runs.append({'kind':'differential','leg':leg,'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'report':report.name,'report_sha256':sha(report) if report.exists() else None,'log':log.name,'log_sha256':sha(log)});write(out/'runs.json',runs);assert r.returncode==0
  print('differential '+leg+' done',flush=True)
 cases=['mce_stream_long_uri','mce_stream_review_long_uri','mce_stream_long_uri_elements','mce_stream_long_uri_skipped','mce_stream_long_uri_tokens','mce_stream_short_uri','mce_stream_benign_worksheet','mce_stream_benign_document','mce_stream_count_worksheet','mce_stream_count_document','mce_extension_long_uri','mce_stream_extension_long_uri','mce_extension_short_uri','mce_stream_extension_short_uri','mce_aliases_low','mce_aliases_high','mce_stream_aliases_low','mce_stream_aliases_high']
 for case in cases:
  samples,warmup=(3,1) if 'long_uri' in case else (9,2)
  for index,leg in enumerate(['before','after','after','before','before','after']):
   prefix=f'{case}-{index}-{leg}';report=out/(prefix+'.json');log=out/(prefix+'.log');rss=out/(prefix+'.rss');binary=Path(binaries[leg]['path']);assert sha(binary)==binaries[leg]['sha256']
   cmd=['/usr/bin/time','-f','%M','-o',str(rss),'taskset','-c','12',str(binary),'adversarial','--case',case,'--samples',str(samples),'--warmup',str(warmup),'--json',str(report)];start=time.time()
   with log.open('w') as f:r=subprocess.run(cmd,cwd=AFTER,env=env,stdout=f,stderr=subprocess.STDOUT)
   runs.append({'kind':'adversarial','leg':leg,'case':case,'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'report':report.name,'report_sha256':sha(report) if report.exists() else None,'log':log.name,'log_sha256':sha(log),'rss':rss.name,'rss_sha256':sha(rss)});write(out/'runs.json',runs);assert r.returncode==0
   print(prefix+' done',flush=True)
 for leg in ROOTS:assert source(ROOTS[leg],REFS[leg])==inputs[leg]
 assert fixtures()==fixture_files
 assert {n:sha(P/'probe-src'/n) for n in probe_inputs}==probe_inputs
 write(out/'complete.json',{'source_unchanged':True,'fixtures_unchanged':True,'serial':True,'runs':len(runs)})
