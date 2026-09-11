"""Build or capture the separately bound public XLSX guard, serially."""
import argparse,datetime,json,os,shutil,subprocess,time
from run import HERE,REPO,SCRATCH,sha,sources,write
PROBE=HERE/'guard-probe'
CANONICAL=['tools/perf-baseline/src/allocation_metrics.rs','tools/perf-baseline/src/bin/support/counting_allocator.rs']
def probe_sources():
    files=sorted([PROBE/'Cargo.toml',PROBE/'Cargo.lock',*PROBE.rglob('*.rs')])
    return {str(p.relative_to(REPO)):sha(p) for p in files} | {p:sha(REPO/p) for p in CANONICAL}
def main():
    parser=argparse.ArgumentParser();parser.add_argument('stage',choices=['before','after']);parser.add_argument('lane',choices=['lock','build','build-allocator','preflight','pilot','r1','r2','allocator-r1','allocator-r2']);args=parser.parse_args()
    stage,lane=args.stage,args.lane;directory=HERE/stage
    env=dict(os.environ,CARGO_TARGET_DIR=str(SCRATCH/'target'),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_RELEASE_DEBUG='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',TMPDIR=str(SCRATCH))
    if lane=='lock':
        assert stage=='before' and not (PROBE/'Cargo.lock').exists()
        command=['cargo','generate-lockfile','--offline','--manifest-path',str(PROBE/'Cargo.toml')]
        with (directory/'guard-lock.log').open('x') as out:r=subprocess.run(command,cwd=REPO,env=env,stdout=out,stderr=subprocess.STDOUT)
        assert r.returncode==0
        write(directory/'guard-lock-receipt.json',{'command':command,'exit_code':0,'lock_sha256':sha(PROBE/'Cargo.lock'),'log_sha256':sha(directory/'guard-lock.log')});print(stage,lane,0,flush=True);return
    build=lane.startswith('build');allocator=lane=='build-allocator' or lane.startswith('allocator')
    manifest=HERE/'source-manifest.json' if stage=='after' else directory/'source-manifest.json'
    current=sources();guard=probe_sources()
    binary=SCRATCH/stage/('xlsx-guard-alloc' if allocator else 'xlsx-guard')
    receipt_name='guard-'+lane
    if build:
        assert current==json.loads(manifest.read_text())
        sourcefile=HERE/'guard-source-manifest.json'
        if sourcefile.exists():assert guard==json.loads(sourcefile.read_text())
        else:write(sourcefile,guard)
        command=['cargo','build','--release','--locked','--offline','--manifest-path',str(PROBE/'Cargo.toml'),'--bin','litchi-xlsx-commit-guard']
        if allocator:command+=['--features','allocator-metrics']
        artifacts=[]
    else:
        receipt=directory/('guard-build-allocator-receipt.json' if allocator else 'guard-build-receipt.json')
        identity=json.loads(receipt.read_text());assert identity['exit_code']==0 and sha(binary)==identity['binary_sha256'];assert guard==json.loads((HERE/'guard-source-manifest.json').read_text())
        samples,warmup=(10,1) if allocator else ((1,0) if lane=='preflight' else ((20,2) if lane=='pilot' else (100,3)))
        report=directory/(receipt_name+'-report.json');assert not report.exists()
        command=[str(binary),'--samples',str(samples),'--warmup',str(warmup),'--shape','tiny,medium,dense-wide','--json',str(report)];artifacts=[report]
    command=['/usr/bin/time','-v',*command] if build else ['/usr/bin/time','-v','taskset','-c','2',*command]
    log=directory/(receipt_name+'.log');start=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
    with log.open('x') as out:r=subprocess.run(command,cwd=REPO,env=env,stdout=out,stderr=subprocess.STDOUT)
    unchanged=sources()==current and probe_sources()==guard
    if build and r.returncode==0:shutil.copy2(SCRATCH/'target/release/litchi-xlsx-commit-guard',binary)
    artifacts.append(log)
    result={'command':command,'started_utc':start,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':unchanged,'source_manifest_sha256':sha(manifest) if build else identity['source_manifest_sha256'],'guard_manifest_sha256':sha(HERE/'guard-source-manifest.json'),'artifacts':{p.name:sha(p) for p in artifacts if p.exists()},'scope':'Separate public guard; per-scenario commit or first-cell clock excludes setup, warming, oracles and result drop; allocator timing/RSS excluded.'}
    if binary.exists():result['binary_sha256']=sha(binary)
    if build:result['environment']={k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_DEBUG','CARGO_PROFILE_DEV_DEBUG','CARGO_PROFILE_TEST_DEBUG','TMPDIR']}
    write(directory/(receipt_name+'-receipt.json'),result);print(stage,lane,r.returncode,flush=True);assert r.returncode==0 and unchanged
if __name__=='__main__':main()
