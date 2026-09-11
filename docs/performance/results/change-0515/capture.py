"""Source-bound current XLSX attribution; no cross-build performance claim."""
import datetime,json,os,subprocess,time,argparse
from run import HERE,REPO,SCRATCH,BINARY,sha,sources,write

def capture(lane):
    build=json.loads((HERE/'build-receipt.json').read_text())
    assert build['exit_code']==0 and sha(BINARY)==build['binary_sha256']
    manifest=json.loads((HERE/'source-manifest.json').read_text());assert sources()==manifest
    samples,warmup=(3,0) if lane.startswith(('commit-', 'compact-')) else ((1,0) if lane=='preflight' else (30,3))
    cases='xlsx_one_percent_commit' if lane.startswith(('commit-', 'compact-')) else 'xlsx_one_cell_commit,xlsx_one_percent_commit,xlsx_one_cell_commit_save,xlsx_one_percent_commit_save'
    shapes='dense-wide' if lane.startswith(('commit-', 'compact-')) else 'tiny,medium,dense-wide'
    report=HERE/(lane+'-report.json');catalog=HERE/(lane+'-catalog.json')
    command=[str(BINARY),'--case',cases,'--xlsx-shape',shapes,'--samples',str(samples),'--warmup',str(warmup),'--json',str(report),'--corpus-manifest',str(catalog)]
    scope='Native existing per-case timer; setup, generator, expected-output generation, verification and drops excluded according to runner'
    if lane.startswith(('commit-', 'compact-')):
        symbols=json.loads((HERE/'profile-symbols.json').read_text())
        compaction = lane.startswith('compact-')
        symbol = symbols['compaction'] if compaction else symbols['commit']
        profile_args=['valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect='+symbol,'--zero-before='+symbols['runner']]
        if not compaction:
            profile_args.append('--separate-callers=3')
        command=[*profile_args,'--callgrind-out-file='+str(HERE/(lane+'.out')),*command]
        scope=('Changed worksheet compaction bodies only; 6 measured calls for 2 changed sheets x 3 commits, reset at runner entry. Grid validation, setup and readback excluded.' if compaction else 'Three commit bodies only, reset at runner entry; three caller contexts distinguish changed-output parser from source Store parser. Setup, readback, save and drops excluded.')
    command=['/usr/bin/time','-v','taskset','-c','2',*command]
    start=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
    for path in (report,catalog,HERE/(lane+'-receipt.json'),HERE/(lane+'.out')):
        assert not path.exists(), f'refusing overwrite: {path}'
    with (HERE/(lane+'.log')).open('x') as out:r=subprocess.run(command,cwd=REPO,stdout=out,stderr=subprocess.STDOUT)
    artifacts=[report,catalog,HERE/(lane+'.log')]
    if lane.startswith(('commit-', 'compact-')):artifacts.append(HERE/(lane+'.out'))
    receipt={'command':command,'started_utc':start,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':sources()==manifest,'source_manifest_sha256':sha(HERE/'source-manifest.json'),'binary_sha256':sha(BINARY),'scope':scope,'artifacts':{p.name:sha(p) for p in artifacts if p.exists()}}
    write(HERE/(lane+'-receipt.json'),receipt);print(lane,r.returncode,flush=True);assert r.returncode==0 and receipt['source_unchanged']

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('lane',choices=['preflight','normal-r1','normal-r2','commit-r1','commit-r2','compact-r1','compact-r2']);capture(parser.parse_args().lane)
