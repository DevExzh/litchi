"""Build the unchanged isolated allocator harness for one frozen role."""
import datetime,json,os,shutil,subprocess,sys,time
from run import HERE,REPO,SCRATCH,sha,sources,write
stage=sys.argv[1];directory=HERE/stage
manifest=directory/'source-manifest.json' if stage=='before' else HERE/'source-manifest.json'
current=sources();assert current==json.loads(manifest.read_text())
env=dict(os.environ,CARGO_TARGET_DIR=str(SCRATCH/'target'),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_RELEASE_DEBUG='0',TMPDIR=str(SCRATCH))
cmd=['cargo','build','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--features','allocator-metrics','--bin','litchi-perf-baseline-alloc']
start=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
with (directory/'allocator-build.log').open('x') as out:result=subprocess.run(cmd,cwd=REPO,env=env,stdout=out,stderr=subprocess.STDOUT)
assert result.returncode==0 and current==sources()
binary=SCRATCH/stage/'litchi-perf-baseline-alloc';shutil.copy2(SCRATCH/'target/release/litchi-perf-baseline-alloc',binary)
write(directory/'allocator-build-receipt.json',{'command':cmd,'started_utc':start,'elapsed_seconds':time.monotonic()-tick,'exit_code':0,'source_unchanged':True,'source_manifest_sha256':sha(manifest),'binary_sha256':sha(binary),'log_sha256':sha(directory/'allocator-build.log'),'environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_DEBUG','TMPDIR']}})
print(stage,'allocator built',flush=True)
