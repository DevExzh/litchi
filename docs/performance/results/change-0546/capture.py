"""Serial capture, run from repository root. Refuses to replace evidence."""
from pathlib import Path
import datetime, hashlib, json, subprocess, sys
B=Path('docs/performance/results/change-0546')
T=Path('/home/zhuhe/litchi-goal-0546-target')
M=T/'scanner/Cargo.toml'
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()
def run(name,argv,binary=None):
    receipt=B/(name+'.receipt.json');stdout=B/(name+'.stdout');stderr=B/(name+'.stderr')
    assert not any(p.exists() for p in [receipt,stdout,stderr]),name
    r={'argv':[str(a) for a in argv],'cwd':str(Path.cwd()),'started':now(),'freeze_sha256':sha(B/'freeze.json')}
    if binary:r['binary_sha256']=sha(binary)
    with stdout.open('wb') as o,stderr.open('wb') as e:result=subprocess.run(r['argv'],stdout=o,stderr=e)
    r.update(ended=now(),exit_code=result.returncode,stdout_sha256=sha(stdout),stderr_sha256=sha(stderr))
    if binary:assert sha(binary)==r['binary_sha256']
    receipt.write_text(json.dumps(r,indent=2)+'\n')
    print(name,result.returncode,flush=True)
    assert result.returncode==0,stderr.read_text()
if sys.argv[1]=='quality':
    for name,args in [('fmt',['fmt','--check']),('test',['test','--offline','--locked']),('clippy',['clippy','--offline','--locked','--all-targets']),('build',['build','--offline','--locked','--release'])]:
        command=['cargo',*args,'--manifest-path',str(M)]
        if name=='clippy':command+=['--','-D','warnings']
        run(name,command)
elif sys.argv[1]=='native':
    binary=T/'scanner/target/release/xlsx-scanner-diagnostic'
    freeze=json.loads((B/'freeze.json').read_text());fixtures=json.loads((B/'fixtures.json').read_text())
    for variant_repeat in freeze['order']:
        variant=variant_repeat.rsplit('-',1)[0]
        for fixture in freeze['fixtures']:
            name=f'native-{variant_repeat}-{fixture}';source=T/'fixtures'/fixtures[fixture]['file']
            assert sha(source)==fixtures[fixture]['sha256']
            run(name,['taskset','-c','2',str(binary),variant,str(source),str(B/(name+'.json'))],binary)
    run('assembly',['objdump','-d','-C',str(binary)],binary)
else:raise ValueError(sys.argv[1])
