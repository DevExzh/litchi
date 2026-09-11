"""Source-bound harness gates, invoked serially by the coordinator."""
import argparse,datetime,json,os,subprocess,time
from run import HERE,REPO,SCRATCH,sha,sources,write
common=['--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml']
commands={
 'tests':['cargo','test',*common,'--features','allocator-metrics','--lib'],
 'fmt':['cargo','fmt','--all','--','--check'],
 'clippy':['cargo','clippy',*common,'--features','allocator-metrics','--lib','--bin','litchi-perf-baseline','--bin','litchi-perf-baseline-alloc','--','-D','warnings'],
 'rustdoc':['cargo','doc',*common,'--features','allocator-metrics','--lib','--no-deps'],
 'boundaries':['python3','-B','tools/check_crate_boundaries.py'],
 'claims':['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--evidence-root','.','--mode','strict'],
}
p=argparse.ArgumentParser();p.add_argument('lane',choices=commands);a=p.parse_args();cmd=commands[a.lane]
manifest=sources();assert manifest==json.loads((HERE/'source-manifest.json').read_text())
env=dict(os.environ,CARGO_TARGET_DIR=str(SCRATCH/'target'),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_RELEASE_DEBUG='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',TMPDIR=str(SCRATCH),RUSTDOCFLAGS='-D warnings')
start=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
with (HERE/(a.lane+'.log')).open('x') as out:r=subprocess.run(cmd,cwd=REPO,env=env,stdout=out,stderr=subprocess.STDOUT)
unchanged=sources()==manifest
write(HERE/(a.lane+'-receipt.json'),{'command':cmd,'started_utc':start,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':unchanged,'source_manifest_sha256':sha(HERE/'source-manifest.json'),'log_sha256':sha(HERE/(a.lane+'.log')),'environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_DEBUG','CARGO_PROFILE_DEV_DEBUG','CARGO_PROFILE_TEST_DEBUG','TMPDIR','RUSTDOCFLAGS']}})
print(a.lane,r.returncode,flush=True);assert r.returncode==0 and unchanged
