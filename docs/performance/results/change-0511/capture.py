"""Serial source-bound CFB/XLS native and scoped diagnostic capture."""
import datetime,json,subprocess,sys,time
from pathlib import Path
from run import HERE,REPO,SCRATCH,sha,sources,write
XLS='xls_semantic_open,xls_eager_open_list_worksheets,xls_eager_open_one_cell,xls_source_backed_open,xls_source_backed_open_list_worksheets,xls_source_backed_open_one_cell,xls_owned_source_open,xls_owned_source_open_list_worksheets,xls_owned_source_open_one_cell'
def run(stage,name,kind='xls'):
    directory=HERE/stage;directory.mkdir(exist_ok=True)
    binary=SCRATCH/stage/('litchi-perf-baseline-alloc' if kind=='allocator' else 'litchi-perf-baseline')
    build_dir=HERE/'before' if stage=='before' else HERE
    build=json.loads(((directory/'allocator-build-receipt.json') if kind=='allocator' else (build_dir/'build-receipt.json')).read_text())
    assert sha(binary)==build['binary_sha256']
    current=sources()
    count=5 if kind in {'callgrind','cfb-profile'} else (1 if kind=='preflight' else (30 if kind=='allocator' else 1000))
    cases='xls_owned_source_open_one_cell' if kind in {'callgrind','preflight','hardware'} else ('cfb_open' if kind in {'guard','cfb-profile'} else XLS)
    command=[str(binary),'--case',cases,'--samples',str(count),'--warmup','0' if kind in {'callgrind','preflight','hardware','cfb-profile'} else ('3' if kind=='allocator' else '20'),'--json',str(directory/f'{name}-report.json'),'--corpus-manifest',str(directory/f'{name}-catalog.json')]
    if kind=='guard':command += ['--shape','tiny,few-large','--payload','incompressible']
    if kind=='cfb-profile':command += ['--shape','few-large','--payload','incompressible']
    if kind=='cfb-profile':
        command=['valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect=*run_cfb_open','--callgrind-out-file='+str(directory/'cfb-profile.out'),*command]
    elif kind=='callgrind':
        command=['valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect=*SourceBackedWorkbook::from_read_at_with_limits','--callgrind-out-file='+str(directory/'callgrind.out'),*command]
    elif kind=='hardware':
        command=['taskset','-c','2','perf','stat','-x',',','-o',str(directory/f'{name}.csv'),'-e','{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations','--',*command]
    else:command=['taskset','-c','2',*command]
    tick=time.monotonic();started=datetime.datetime.now(datetime.timezone.utc).isoformat()
    with (directory/f'{name}.log').open('x') as out:
        result=subprocess.run(['/usr/bin/time','-v',*command],cwd=REPO,stdout=out,stderr=subprocess.STDOUT)
    assert result.returncode==0,(directory/f'{name}.log').read_text()[-3000:]
    assert current==sources() and sha(binary)==build['binary_sha256']
    names=[f'{name}.log',f'{name}-report.json',f'{name}-catalog.json']+(['callgrind.out'] if kind=='callgrind' else (['cfb-profile.out'] if kind=='cfb-profile' else ([f'{name}.csv'] if kind=='hardware' else [])))
    write(directory/f'{name}-receipt.json',{'command':command,'started_utc':started,'elapsed_seconds':time.monotonic()-tick,'exit_code':0,'binary_sha256':sha(binary),'source_manifest_sha256':sha(build_dir/'source-manifest.json'),'source_unchanged':True,'artifacts':{n:sha(directory/n) for n in names},'scope':'CFB guard runner including five opens, timers, file-size oracles, drops and result construction; fixture generation excluded' if kind=='cfb-profile' else 'SourceBackedWorkbook::from_read_at_with_limits only; excludes selected-cell query and setup; simulated instruction references' if kind=='callgrind' else ('whole-child hardware counters including setup, copies, queries and oracles; instrumented timings excluded' if kind=='hardware' else 'instrumented operation allocation deltas; timings excluded from native comparisons' if kind=='allocator' else 'existing matched operation clocks; whole-child RSS includes setup, copies and oracles')})
    print(stage,name,'complete',flush=True)
if __name__=='__main__':
    if sys.argv[1]=='hardware':
        for stage,repeat in [('before','r1'),('after','r1'),('after','r2'),('before','r2')]:run(stage,f'hardware-{repeat}','hardware')
    elif sys.argv[1]=='allocation':
        for stage,repeat in [('before','r1'),('after','r1'),('after','r2'),('before','r2')]:run(stage,f'allocator-{repeat}','allocator')
    elif sys.argv[1]=='timing':
        for kind in ['xls','guard']:
            for stage,repeat in [('before','r1'),('after','r1'),('after','r2'),('before','r2')]:run(stage,f'{kind}-{repeat}',kind)
    else:run(sys.argv[1],sys.argv[2],sys.argv[2])
