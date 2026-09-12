"""Build and freeze a source-bound 0516 control or candidate executable."""
import argparse, datetime, hashlib, json, os, shutil, subprocess, time
from pathlib import Path
HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SCRATCH = Path('/tmp/litchi-goal-0516')

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def sources():
    paths = subprocess.check_output(['git','ls-files','--cached','--others','--exclude-standard','-z','crates','tools/perf-baseline','Cargo.toml','Cargo.lock','rust-toolchain.toml','.cargo'],cwd=REPO).decode().split('\0')
    return {p:sha(REPO/p) for p in sorted(paths) if p and Path(p).suffix in {'.rs','.toml','.lock'}}

def write(path, value):
    with path.open('x') as out:
        json.dump(value,out,indent=2);out.write('\n')

def main():
    parser=argparse.ArgumentParser();parser.add_argument('stage',choices=['before','after']);parser.add_argument('kind',choices=['normal','allocator']);a=parser.parse_args()
    directory=HERE/a.stage;directory.mkdir(exist_ok=True)
    binary_dir=SCRATCH/a.stage;binary_dir.mkdir(exist_ok=True)
    current=sources();manifest=directory/'source-manifest.json'
    if manifest.exists():assert current==json.loads(manifest.read_text())
    else:write(manifest,current)
    allocator=a.kind=='allocator';name='litchi-perf-baseline-alloc' if allocator else 'litchi-perf-baseline'
    lane='allocator-build' if allocator else 'build'
    command=['cargo','build','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--bin',name]
    if allocator:command+=['--features','allocator-metrics']
    env=dict(os.environ,CARGO_TARGET_DIR=str(SCRATCH/'target'),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_RELEASE_DEBUG='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',TMPDIR=str(SCRATCH))
    started=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
    with (directory/(lane+'.log')).open('x') as out:r=subprocess.run(['/usr/bin/time','-v',*command],cwd=REPO,env=env,stdout=out,stderr=subprocess.STDOUT)
    unchanged=sources()==current
    receipt={'stage':a.stage,'kind':a.kind,'command':command,'started_utc':started,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':unchanged,'source_manifest_sha256':sha(manifest),'log_sha256':sha(directory/(lane+'.log')),'environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_DEBUG','CARGO_PROFILE_DEV_DEBUG','CARGO_PROFILE_TEST_DEBUG','TMPDIR']}}
    if r.returncode==0:
        binary=binary_dir/name;assert not binary.exists();shutil.copy2(SCRATCH/'target/release'/name,binary);receipt['binary_sha256']=sha(binary)
    write(directory/(lane+'-receipt.json'),receipt)
    print(a.stage,a.kind,r.returncode,flush=True);assert r.returncode==0 and unchanged
if __name__=='__main__':main()
