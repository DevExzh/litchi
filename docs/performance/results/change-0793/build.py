"""Root-only serial builds of diagnostic variants, with immutable receipts."""
import os,shutil,subprocess,time
import custody as c
assert not c.TARGET.exists()
out=c.P/'build';assert not out.exists();out.mkdir()
manifest=c.P/'probe-src/Cargo.toml';manifest.write_text((c.P/'probe-src/Cargo.toml.template').read_text().replace('@SRC@',str(c.ROOT)))
source=c.source();c.write(out/'source.json',source)
assert source['files']==c.read(c.P/'../change-0792/build-after/source.json')['files']
inputs=['plan.json','origin.json','host.json','architecture-inputs.json','inheritance.json','build.py','capture.py','decode.py']
c.write(out/'frozen-inputs.json',{n:c.sha(c.P/n) for n in inputs})
probe={str(f.relative_to(c.P)):c.sha(f) for f in (c.P/'probe-src').rglob('*') if f.is_file()};c.write(out/'probe.json',probe)
variants=[('control',[]),('profile',['capture-profile']),('allocation',['allocator-metrics']),('profile-allocation',['allocator-metrics','capture-profile'])]
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
rows=[];binaries={}
for name,features in variants:
 cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(manifest)]
 if features:cmd+=['--features',','.join(features)]
 log=out/f'{name}.log';start=time.time()
 with log.open('w') as stream:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=stream,stderr=subprocess.STDOUT)
 rows.append({'name':name,'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'log':c.artifact(log)});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 assert c.source()==source and all(c.sha(c.P/n)==h for n,h in probe.items())
 dest=c.TARGET/name;shutil.copy2(c.TARGET/'release/namespace-uri-probe',dest);binaries[name]=c.artifact(dest)
 print(name,'built',flush=True)
c.write(out/'receipt.json',{'source':c.artifact(out/'source.json'),'probe':probe,'rows':rows,'binaries':binaries,'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTUP_TOOLCHAIN']}})
