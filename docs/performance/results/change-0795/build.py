"""Root-only source-bound ordinary and frame-pointer diagnostic builds."""
import os,sys,subprocess,shutil,time
import custody as c
leg=sys.argv[1];assert leg in ['before','after']
out=c.P/f'build-{leg}';assert not out.exists();out.mkdir()
source=c.source();c.write(out/'source.json',source)
expected=c.read(c.P/'baseline-source.json')['files'].copy()
if leg=='after':
 for r in c.read(c.P.parent/'change-0794/candidate/manifest.json')['files']:expected[r['production']]=r['after_sha256']
assert source['files']==expected
assert not os.environ.get('RUSTFLAGS')
frozen={n:c.sha(c.P/n) for n in ['plan.json','origin.json','inheritance.json','architecture-inputs.json','host.json','build.py','capture.py','custody.py']};c.write(out/'frozen.json',frozen)
probe={str(f.relative_to(c.P/'probe-src')):c.sha(f) for f in (c.P/'probe-src').rglob('*') if f.is_file() and f.name!='Cargo.toml'}
assert probe==c.read(c.P/'inheritance.json')['probe']
rows=[];binaries={}
for kind,flags in [('profile',''),('fp','-C force-frame-pointers=yes')]:
 env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','RUSTFLAGS':flags}
 cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(c.P/'probe-src/Cargo.toml'),'--features','capture-profile']
 log=out/f'{kind}.log';started=time.time()
 with log.open('w') as stream:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=stream,stderr=subprocess.STDOUT)
 rows.append({'kind':kind,'command':cmd,'started':started,'ended':time.time(),'exit_code':r.returncode,'log':c.artifact(log),'environment':{n:env[n] for n in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS']}});c.write(out/'commands.json',rows)
 assert r.returncode==0 and c.source()==source
 binary=c.TARGET/f'{leg}-{kind}';assert not binary.exists();shutil.copy2(c.TARGET/'release/namespace-uri-probe',binary);binaries[kind]=c.artifact(binary)
 print(leg,kind,'PASS',flush=True)
c.write(out/'build.json',{'source':c.artifact(out/'source.json'),'frozen':frozen,'probe':probe,'rows':rows,'binaries':binaries})
