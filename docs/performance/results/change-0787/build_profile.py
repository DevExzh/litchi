"""Root-only current-baseline profile build; no production modifications."""
import os,subprocess,shutil,time
import custody as c
out=c.P/'profile-build';assert not out.exists();out.mkdir()
manifest=c.P/'profile-src/Cargo.toml';manifest.write_text((c.P/'profile-src/Cargo.toml.template').read_text().replace('@SRC@',str(c.ROOT)))
frozen=c.source();assert frozen==c.read(c.P/'build-before/source.json')['production'];c.write(out/'source.json',frozen)
probe={str(p.relative_to(c.P)):c.sha(p) for p in (c.P/'profile-src').rglob('*') if p.is_file()};c.write(out/'probe.json',probe)
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
cmd=['cargo','build','--offline','--locked','--release','--manifest-path',str(manifest)];log=out/'build.log';start=time.time()
with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
receipt={'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'log':c.artifact(log),'source':c.artifact(out/'source.json'),'probe':c.artifact(out/'probe.json'),'plan_sha256':c.sha(c.P/'profile-plan.json'),'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS']}}
c.write(out/'receipt.json',receipt);assert r.returncode==0,log;assert c.source()==frozen
assert probe=={str(p.relative_to(c.P)):c.sha(p) for p in (c.P/'profile-src').rglob('*') if p.is_file()}
dest=c.TARGET/'baseline-profile';assert not dest.exists();shutil.copy2(c.TARGET/'release/cached-part-profile',dest);receipt['binary']=c.artifact(dest)
cmd=['nm','-C',str(dest)];symbols=subprocess.check_output(cmd,text=True);matches=[line for line in symbols.splitlines() if line.endswith('cached_part_profile::cached_part_region_0787')];assert len(matches)==1,matches
(out/'owner-symbol.txt').write_text(matches[0]+'\n');receipt['owner']='cached_part_profile::cached_part_region_0787';receipt['owner_symbol']=c.artifact(out/'owner-symbol.txt');c.write(out/'receipt.json',receipt);print('baseline profile built and owner symbol qualified',flush=True)
