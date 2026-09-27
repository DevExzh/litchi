"""Root-only cross-format harness build with exact source/tool custody."""
import os,shutil,subprocess,sys,time
import custody as c
leg=sys.argv[1];assert leg in ['before','after']
out=c.P/f'cross-build-{leg}';assert not out.exists();out.mkdir()
source=c.source();c.write(out/'source.json',source)
if leg=='before':assert source==c.read(c.P/'build-before/source.json')
files=subprocess.check_output(['git','ls-files','-z','--','tools/perf-baseline'],cwd=c.ROOT).decode().split('\0')
inventory={n:c.sha(c.ROOT/n) for n in files if n};c.write(out/'tools.json',inventory)
assert c.sha(c.ROOT/'tools/perf-baseline/Cargo.lock')==c.sha(c.P/'cross-Cargo.lock')
c.write(out/'frozen-inputs.json',{n:c.sha(c.P/n) for n in ['cross-plan.json','cross_build.py','cross_capture.py','cross-Cargo.lock']})
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET/'cross'),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(c.ROOT/'tools/perf-baseline/Cargo.toml'),'--bin','litchi-perf-baseline']
log=out/'build.log';started=time.time()
with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
row={'command':cmd,'exit_code':r.returncode,'started':started,'ended':time.time(),'log':c.artifact(log),'source':c.artifact(out/'source.json'),'tools':c.artifact(out/'tools.json'),'lock':c.artifact(c.P/'cross-Cargo.lock'),'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTUP_TOOLCHAIN']}}
c.write(out/'receipt.json',row);assert r.returncode==0
assert c.source()==source
assert inventory=={n:c.sha(c.ROOT/n) for n in inventory}
binary=c.TARGET/f'{leg}-cross';assert not binary.exists();shutil.copy2(c.TARGET/'cross/release/litchi-perf-baseline',binary)
row['binary']=c.artifact(binary);c.write(out/'receipt.json',row)
print(leg,'cross-format native built',flush=True)
