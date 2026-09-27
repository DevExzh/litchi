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
out=c.P/f'build-{leg}'
assert not out.exists()
out.mkdir()
manifest=c.P/'probe-src/Cargo.toml'
manifest.write_text((c.P/'probe-src/Cargo.toml.template').read_text().replace('@SRC@',str(c.ROOT)))
frozen=c.source();c.write(out/'source.json',frozen)
probe_files={str(p.relative_to(c.P)):c.sha(p) for p in (c.P/'probe-src').rglob('*') if p.is_file() and p.name not in ('Cargo.lock','Cargo.toml')}
c.write(out/'probe.json',probe_files)
env=os.environ|{'CARGO_TARGET_DIR':str(c.TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0'}
rows=[]; binaries={}
for name,features in [('native',[]),('allocation',['--features','allocator-metrics'])]:
 command=['cargo','build','--offline','--release','--manifest-path',str(manifest),*features]
 if (manifest.parent/'Cargo.lock').is_file():command.append('--locked')
 log=out/f'{name}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(command,cwd=c.ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':command,'exit_code':r.returncode,'log':c.artifact(log),'started':start,'ended':time.time()});c.write(out/'commands.json',rows)
 assert r.returncode==0,log
 assert c.source()==frozen
 source=c.TARGET/'release/xlsx-allocation-probe';dest=c.TARGET/f'{leg}-{name}'
 assert not dest.exists();shutil.copy2(source,dest);binaries[name]=c.artifact(dest)
 print(leg,name,'built',flush=True)
c.write(out/'build.json',{'source':c.artifact(out/'source.json'),'probe':probe_files,'binaries':binaries,'lock':c.artifact(manifest.parent/'Cargo.lock'),'rows':rows,'environment':{k:env.get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTUP_TOOLCHAIN']}})
