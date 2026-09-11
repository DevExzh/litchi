"""Source-bound current XLSX attribution; no cross-build performance claim."""
import datetime,json,os,subprocess,time,argparse
from run import HERE,REPO,SCRATCH,BINARY,sha,sources,write

def capture(lane):
    build=json.loads((HERE/'build-receipt.json').read_text())
    assert build['exit_code']==0 and sha(BINARY)==build['binary_sha256']
    manifest=json.loads((HERE/'source-manifest.json').read_text());assert sources()==manifest
    samples,warmup=(3,0) if lane.startswith('profile') else ((1,0) if lane=='preflight' else (30,3))
    cases='xlsx_one_percent_commit' if lane.startswith('profile') or lane=='hardware' else 'xlsx_one_cell_commit,xlsx_one_percent_commit,xlsx_one_cell_commit_save,xlsx_one_percent_commit_save'
    shapes='dense-wide' if lane.startswith('profile') or lane=='hardware' else 'tiny,medium,dense-wide'
    report=HERE/(lane+'-report.json');catalog=HERE/(lane+'-catalog.json')
    command=[str(BINARY),'--case',cases,'--xlsx-shape',shapes,'--samples',str(samples),'--warmup',str(warmup),'--json',str(report),'--corpus-manifest',str(catalog)]
    scope='Native existing per-case timer; setup, generator, expected-output generation, verification and drops excluded according to runner'
    if lane.startswith('profile'):
        symbols=json.loads((HERE/'profile-symbols.json').read_text())
        command=['valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect='+symbols['commit'],'--zero-before='+symbols['runner'],'--callgrind-out-file='+str(HERE/(lane+'.out')),*command]
        scope='Commit function only, reset at update-commit runner entry to exclude generator commit; 3 measured commits, no warmups; setup, final readback, save and drop excluded'
    elif lane=='hardware':
        command=['perf','stat','-x,','-o',str(HERE/(lane+'.csv')),'-e','{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations',*command]
        scope='Whole child including fixture generation, opening, staging, 3 warmups, timed commits, final readback and teardown; not operation-local counters'
    command=['/usr/bin/time','-v','taskset','-c','2',*command]
    start=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
    with (HERE/(lane+'.log')).open('x') as out:r=subprocess.run(command,cwd=REPO,stdout=out,stderr=subprocess.STDOUT)
    artifacts=[report,catalog,HERE/(lane+'.log')]
    if lane.startswith('profile'):artifacts.append(HERE/(lane+'.out'))
    if lane=='hardware':artifacts.append(HERE/(lane+'.csv'))
    receipt={'command':command,'started_utc':start,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':sources()==manifest,'source_manifest_sha256':sha(HERE/'source-manifest.json'),'binary_sha256':sha(BINARY),'scope':scope,'artifacts':{p.name:sha(p) for p in artifacts if p.exists()}}
    write(HERE/(lane+'-receipt.json'),receipt);print(lane,r.returncode,flush=True);assert r.returncode==0 and receipt['source_unchanged']

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('lane',choices=['preflight','normal-r1','normal-r2','profile-r1','profile-r2','hardware']);capture(parser.parse_args().lane)
