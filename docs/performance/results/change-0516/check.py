"""Source-bound harness gates, invoked serially by the coordinator."""
import argparse,datetime,json,os,subprocess,time
from run import HERE,REPO,SCRATCH,sha,sources,write
common=['--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml']
commands={
 'tests':['cargo','test',*common,'--features','allocator-metrics','--lib'],
 'xlsx-unit':['cargo','test','--locked','-p','litchi-xlsx','--lib'],
 'xlsx-default':['cargo','test','--locked','-p','litchi-xlsx'],
 'xlsx-features':['cargo','test','--locked','-p','litchi-xlsx','--all-features'],
 'owner-check':['cargo','check','--locked','-p','litchi-xlsx','--all-features'],
 'workspace-check':['cargo','check','--locked','--workspace','--all-features'],
 'fmt':['cargo','fmt','--all','--','--check'],
 'clippy':['cargo','clippy','--locked','-p','litchi-xlsx','--all-features','--lib','--','-D','warnings'],
 'rustdoc':['cargo','doc','--locked','-p','litchi-xlsx','--all-features','--no-deps'],
 'boundaries':['python3','-B','tools/check_crate_boundaries.py'],
 'claims':['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--evidence-root','.','--mode','strict'],
}

commands['xlsx-source-backed-values']=['cargo','test','--locked','-p','litchi-xlsx','--all-features','--test','source_backed_cell_values']
commands['xlsx-features-retry']=commands['xlsx-features']
commands['xlsx-features-known']=commands['xlsx-features']+['--','--skip','managed_scalar_exact_noop_publishes_without_detaching_source','--skip','managed_multi_sheet_exact_noop_publishes_without_detaching_sources']

commands['xlsx-features-known2']=commands['xlsx-features-known']+['--skip','managed_exact_noop_publication_is_byte_exact_and_releases_budget','--skip','managed_signature_noop_and_changed_protection_contracts_remain_fail_closed']

p=argparse.ArgumentParser();p.add_argument('lane',choices=commands);p.add_argument('--epoch',default='after');a=p.parse_args();cmd=commands[a.lane]
directory=HERE/a.epoch;directory.mkdir(exist_ok=True)
manifest=sources();assert manifest==json.loads((directory/'source-manifest.json').read_text())
env=dict(os.environ,CARGO_TARGET_DIR=str(SCRATCH/'target'),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_RELEASE_DEBUG='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',TMPDIR=str(SCRATCH),RUSTDOCFLAGS='-D warnings')
start=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
with (directory/(a.lane+'.log')).open('x') as out:r=subprocess.run(cmd,cwd=REPO,env=env,stdout=out,stderr=subprocess.STDOUT)
unchanged=sources()==manifest
write(directory/(a.lane+'-receipt.json'),{'command':cmd,'started_utc':start,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':unchanged,'source_manifest_sha256':sha(directory/'source-manifest.json'),'log_sha256':sha(directory/(a.lane+'.log')),'environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_DEBUG','CARGO_PROFILE_DEV_DEBUG','CARGO_PROFILE_TEST_DEBUG','TMPDIR','RUSTDOCFLAGS']}})
print(a.lane,r.returncode,flush=True);assert r.returncode==0 and unchanged
