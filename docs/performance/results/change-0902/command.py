"""Root-only, write-once build and quality command receipt."""
import argparse,datetime,hashlib,json,os,pathlib,subprocess
P=pathlib.Path(__file__).resolve().parent
S=pathlib.Path('/tmp/litchi-0902-src');T=pathlib.Path('/home/zhuhe/code/litchi/target/litchi-0902-build')
def desc(p):
 p=pathlib.Path(p).resolve();return {'path':str(p),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
def write(p,v):
 with p.open('x') as f:json.dump(v,f,indent=2,sort_keys=True);f.write('\n')
a=argparse.ArgumentParser();a.add_argument('label');a.add_argument('argv',nargs=argparse.REMAINDER);args=a.parse_args();argv=args.argv
if argv and argv[0]=='--':argv=argv[1:]
assert argv
base=P/args.label;assert not (P/(args.label+'-start.json')).exists()
env={'PATH':'/home/zhuhe/.cargo/bin:/usr/bin:/bin','HOME':'/home/zhuhe','LC_ALL':'C','LANG':'C','TZ':'UTC','RUSTUP_TOOLCHAIN':'1.95.0','CARGO_BUILD_JOBS':'2','CARGO_TARGET_DIR':str(T),'CARGO_INCREMENTAL':'0','CARGO_NET_OFFLINE':'true','CARGO_PROFILE_DEV_DEBUG':'0','CARGO_PROFILE_TEST_DEBUG':'0','CARGO_PROFILE_RELEASE_OPT_LEVEL':'3','CARGO_PROFILE_RELEASE_DEBUG':'1','CARGO_PROFILE_RELEASE_LTO':'fat','CARGO_PROFILE_RELEASE_CODEGEN_UNITS':'1'}
inputs=[S/'crates/xml-minifier/src/audit.rs',S/'crates/xml-minifier/tests/source_replacement.rs',S/'Cargo.lock',S/'tools/perf-baseline/Cargo.lock']
receipt={'argv':argv,'cwd':str(S),'environment':env,'started_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'helper':desc(__file__),'inputs_before':[desc(f) for f in inputs]};write(P/(args.label+'-start.json'),receipt)
log=P/(args.label+'.log')
with log.open('x') as stream:r=subprocess.run(argv,cwd=S,env=env,stdout=stream,stderr=subprocess.STDOUT)
receipt.update(finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),exit_code=r.returncode,log=desc(log),inputs_after=[desc(f) for f in inputs]);receipt['inputs_unchanged']=receipt['inputs_before']==receipt['inputs_after'];write(P/(args.label+'-run.json'),receipt)
print(args.label,'exit',r.returncode,'source stable',receipt['inputs_unchanged'],flush=True)
if r.returncode or not receipt['inputs_unchanged']:raise SystemExit(r.returncode or 1)
