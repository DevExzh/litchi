"""Validate the authorized cherry-pick in an isolated, owned worktree."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
SOURCE=ROOT.parent/'litchi-worktrees/0913-iwa-cherry-resolution'
TARGET=ROOT/'target/litchi-0913-iwa-build'
PATHS=('crates/litchi-iwa-archive/src/ingress.rs','crates/litchi-iwa-archive/src/lib.rs','crates/litchi-iwa-archive/src/package.rs','crates/litchi-iwa-core/src/archive.rs','crates/litchi-iwa-core/src/archive/ingress.rs','crates/litchi-iwa-core/src/archive/inventory.rs','crates/litchi-iwa-core/src/lib.rs','crates/litchi-iwa-core/src/snappy.rs')
def desc(f): return {'path':str(f.relative_to(ROOT)) if f.is_relative_to(ROOT) else str(f),'bytes':f.stat().st_size,'sha256':hashlib.sha256(f.read_bytes()).hexdigest()}
def new(f,value):
    with f.open('x') as out:json.dump(value,out,indent=2);out.write('\n')
def main():
    assert (SOURCE/'.owned-0913-iwa').read_text()=='iwa-cherry-resolution-0913\n'
    for name in PATHS: assert (SOURCE/name).read_bytes()==(ROOT/name).read_bytes(),name
    assert (SOURCE/'Cargo.lock').read_bytes()==(P/'iwa-Cargo.lock').read_bytes()
    env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',RUSTDOCFLAGS='-D warnings',LC_ALL='C',LANG='C',TZ='UTC')
    for key in ('RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS'):env.pop(key,None)
    packages=['-p','litchi-iwa-core','-p','litchi-iwa-archive']; common=['--locked','--offline','--all-features','-j','2']
    commands=[('tests',['cargo','test',*packages,*common,'--all-targets']),('doctests',['cargo','test',*packages,*common,'--doc']),('clippy',['cargo','clippy',*packages,*common,'--all-targets','--','-D','warnings']),('docs',['cargo','doc',*packages,*common,'--no-deps']),('boundaries',['python3','tools/check_crate_boundaries.py','--explain'])]
    rows=[]
    for kind,argv in commands:
        folder=P/'iwa-quality-v2'/kind;folder.mkdir(parents=True)
        start=datetime.datetime.now(datetime.timezone.utc).isoformat()
        with (folder/'stdout.txt').open('x') as out,(folder/'stderr.txt').open('x') as err:result=subprocess.run(argv,cwd=SOURCE,env=env,stdout=out,stderr=err)
        row={'argv':argv,'cwd':str(SOURCE),'start_utc':start,'end_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':result.returncode,'helper':desc(Path(__file__)),'source':desc(P/'iwa-quality-source.json'),'quality_lock':desc(P/'iwa-quality-lock.json'),'logs':[desc(folder/n) for n in ('stdout.txt','stderr.txt')]}
        new(folder/'command.json',row);rows.append(row);print(kind,result.returncode,flush=True)
        if result.returncode:raise SystemExit(result.returncode)
    new(P/'iwa-quality.json',{'status':'pass','commands':rows})
if __name__=='__main__':main()
