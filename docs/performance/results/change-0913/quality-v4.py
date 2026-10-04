"""Serial correctness and quality gates on the isolated candidate source."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
SOURCE = Path('/tmp/litchi-0913-src')
TARGET = ROOT / 'target/litchi-0913-build'

def desc(path):
    return {'path': str(path.relative_to(ROOT)) if path.is_relative_to(ROOT) else str(path), 'bytes': path.stat().st_size, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}

def main():
    env = dict(os.environ)
    env.update(CARGO_TARGET_DIR=str(TARGET), CARGO_BUILD_JOBS='2', CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_LTO='false', CARGO_PROFILE_RELEASE_CODEGEN_UNITS='16', CARGO_PROFILE_RELEASE_DEBUG='0', LC_ALL='C', LANG='C', TZ='UTC', RUSTDOCFLAGS='-D warnings')
    for key in ('RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS'):
        env.pop(key,None)
    changed = ['crates/soapberry-zip/src/office.rs','crates/litchi-opc/src/payload.rs','crates/litchi-opc/tests/owned_direct_concurrency.rs']
    changed += [str(f.relative_to(SOURCE)) for f in (SOURCE/'crates/soapberry-zip/tests').glob('*owned*')]
    common = ['--locked','--release','-j','2','--features','litchi-opc/sign,litchi-opc/performance-diagnostics']
    packages = ['-p','soapberry-zip','-p','litchi-opc','-p','litchi-pptx','-p','litchi-docx','-p','litchi-xlsx','-p','litchi-xlsb']
    commands = [
        ('fmt',['python3','-B',str(P/'fmt.py')]),
        ('tests',['cargo','test',*packages,*common]),
        ('clippy',['cargo','clippy',*packages,'--all-targets',*common,'--','-D','warnings']),
        ('docs',['cargo','doc',*packages,'--no-deps',*common]),
    ]
    receipts=[]
    for name,argv in commands:
        folder=P/'quality-v4'/name
        folder.mkdir(parents=True)
        start=datetime.datetime.now(datetime.timezone.utc).isoformat()
        with (folder/'stdout.txt').open('x') as out,(folder/'stderr.txt').open('x') as err:
            result=subprocess.run(argv,cwd=SOURCE,env=env,stdout=out,stderr=err)
        receipt={'argv':argv,'cwd':str(SOURCE),'started_utc':start,'ended_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':result.returncode,'helper':desc(Path(__file__)),'source':desc(P/'source-v1.json'),'fixtures':desc(P/'fixtures.json'),'quality_root_lock':desc(P/'quality-lock-v3/Cargo.lock'),'environment_overrides':{k:env[k] for k in ('CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_LTO','CARGO_PROFILE_RELEASE_CODEGEN_UNITS','CARGO_PROFILE_RELEASE_DEBUG','LC_ALL','LANG','TZ','RUSTDOCFLAGS')},'removed_environment_keys':['RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS'],'logs':[desc(folder/n) for n in ('stdout.txt','stderr.txt')]}
        with (folder/'command.json').open('x') as out: json.dump(receipt,out,indent=2);out.write('\n')
        receipts.append(receipt)
        print(name,result.returncode,flush=True)
        if result.returncode: raise SystemExit(result.returncode)
    with (P/'quality.json').open('x') as out: json.dump({'status':'pass','commands':[r['argv'] for r in receipts]},out,indent=2);out.write('\n')

if __name__=='__main__':
    main()
