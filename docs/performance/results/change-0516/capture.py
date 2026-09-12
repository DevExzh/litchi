"""Capture frozen control/candidate XLSX operations without replacing evidence."""
import argparse,datetime,json,subprocess,time
from run import HERE,REPO,SCRATCH,sha,sources,write
CASES='xlsx_one_cell_commit,xlsx_one_percent_commit,xlsx_one_cell_commit_save,xlsx_one_percent_commit_save'

def active_roles(current):
    roles=[role for role in ['before','after'] if (HERE/role/'source-manifest.json').exists() and json.loads((HERE/role/'source-manifest.json').read_text())==current]
    assert roles,'active source does not match a frozen role'
    return roles

def capture(stage,lane):
    directory=HERE/stage;allocator=lane.startswith('allocator')
    build=json.loads((directory/('allocator-build-receipt.json' if allocator else 'build-receipt.json')).read_text());assert build['exit_code']==0
    binary=SCRATCH/stage/('litchi-perf-baseline-alloc' if allocator else 'litchi-perf-baseline');assert sha(binary)==build['binary_sha256']
    current=sources();roles=active_roles(current)
    samples,warmup=(10,1) if allocator else ((3,0) if lane=='profile' else ((1,0) if lane=='preflight' else ((20,2) if lane=='pilot' else (500,5))))
    report=directory/(lane+'-report.json');catalog=directory/(lane+'-catalog.json')
    for path in [report,catalog,directory/(lane+'-receipt.json'),directory/(lane+'.out')]:assert not path.exists(),path
    command=[str(binary),'--case','xlsx_one_percent_commit_save' if lane=='profile' else CASES,'--xlsx-shape','dense-wide' if lane=='profile' else 'tiny,medium,dense-wide','--samples',str(samples),'--warmup',str(warmup),'--json',str(report),'--corpus-manifest',str(catalog)]
    scope='Existing native commit or commit+save clock; setup, expected output, sink reservation, oracles and returned-result drop excluded'
    if lane=='profile':
        command=['valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect=*litchi_perf_baseline::xlsx_commit_save_operation','--callgrind-out-file='+str(directory/'profile.out'),*command]
        scope='Three xlsx_commit_save_operation helper calls only; fixture and expected-output commits/writes excluded; helper returns Commit before caller drop'
    if allocator:scope='Allocation region begins before Instant and ends after elapsed; operation only, no setup/oracles/returned-result drop; instrumented elapsed and RSS excluded'
    command=['/usr/bin/time','-v','taskset','-c','2',*command]
    started=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
    with (directory/(lane+'.log')).open('x') as out:r=subprocess.run(command,cwd=REPO,stdout=out,stderr=subprocess.STDOUT)
    artifacts=[report,catalog,directory/(lane+'.log')]
    if lane=='profile':artifacts.append(directory/'profile.out')
    receipt={'stage':stage,'command':command,'started_utc':started,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':sources()==current,'active_source_roles':roles,'source_manifest_sha256':build['source_manifest_sha256'],'binary_sha256':sha(binary),'scope':scope,'artifacts':{p.name:sha(p) for p in artifacts if p.exists()}}
    write(directory/(lane+'-receipt.json'),receipt);print(stage,lane,r.returncode,flush=True);assert r.returncode==0 and receipt['source_unchanged']
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('stage',choices=['before','after']);p.add_argument('lane',choices=['preflight','pilot','r1','r2','allocator-r1','allocator-r2','profile']);a=p.parse_args();capture(a.stage,a.lane)
