"""Root-only serial probe build with source and binary custody."""
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time
import custody as c

leg=sys.argv[1]
assert leg in ('before','after')
if leg=='before':
 assert subprocess.run(['git','diff','--quiet','HEAD','--','crates','Cargo.toml','clippy.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=c.ROOT).returncode==0, 'baseline production source is modified'
out=c.P/f'build-{leg}'
assert not out.exists()
out.mkdir()
manifest=c.P/'probe-src/Cargo.toml'
manifest.write_text((c.P/'probe-src/Cargo.toml.template').read_text().replace('@SRC@',str(c.ROOT)))
frozen=c.source();c.write(out/'source.json',frozen)
c.write(out/'frozen-inputs.json',{n:c.sha(c.P/n) for n in ['plan.json','adoption-policy.json','build.py','capture.py','architecture-inputs.json','analysis-plan.json','quality.py','probe_quality.py','profile.py','origin.json','inheritance.json','host.json']})
probe_files={str(p.relative_to(c.P)):c.sha(p) for p in (c.P/'probe-src').rglob('*') if p.is_file() and p.name not in ('Cargo.lock','Cargo.toml')}
c.write(out/'probe.json',probe_files)
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
rows=[]; binaries={}
for name,features in [('native',[]),('allocation',['--features','allocator-metrics']),('profile',['--features','capture-profile'])]:
 command=['cargo','build','--offline','--release','--manifest-path',str(manifest),*features]
 if (manifest.parent/'Cargo.lock').is_file():command.append('--locked')
 log=out/f'{name}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(command,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':command,'exit_code':r.returncode,'log':c.artifact(log),'started':start,'ended':time.time()});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 assert c.source()==frozen
 source=c.TARGET/'release/namespace-uri-probe';dest=c.TARGET/f'{leg}-{name}'
 assert not dest.exists();shutil.copy2(source,dest);binaries[name]=c.artifact(dest)
 print(leg,name,'built',flush=True)
c.write(out/'build.json',{'source':c.artifact(out/'source.json'),'probe':probe_files,'binaries':binaries,'lock':c.artifact(manifest.parent/'Cargo.lock'),'rows':rows,'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTUP_TOOLCHAIN']}})
