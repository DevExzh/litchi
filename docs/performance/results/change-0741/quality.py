"""Serial affected-owner checks; preserve every attempted gate and source binding."""
import json,os,subprocess,time
from build import P,ROOT,census,sha
if __name__=='__main__':
    before=census();attempt=0
    while (P/f'quality-{attempt}').exists():attempt+=1
    folder=P/f'quality-{attempt}';folder.mkdir()
    target=ROOT.parent/'litchi-target-0741-quality'
    env=os.environ|{'CARGO_TARGET_DIR':str(target),'CARGO_BUILD_JOBS':'2','RUSTDOCFLAGS':'-D warnings','PYTHONDONTWRITEBYTECODE':'1'}
    owners=['-p','litchi-opc','-p','litchi-pptx']
    common=[*owners,'--release','--offline','--locked']
    commands=[['cargo','fmt',*owners,'--','--check'],['cargo','check',*common,'--all-features'],['cargo','test',*common,'--all-targets'],['cargo','clippy',*common,'--all-targets','--','-D','warnings'],['cargo','test',*common,'--doc'],['cargo','doc',*common,'--no-deps'],['python3','-B','tools/check_crate_boundaries.py']]
    rows=[]
    for i,cmd in enumerate(commands):
        start=time.time();log=folder/f'{i:02d}.log'
        with log.open('w') as stream:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=stream,stderr=subprocess.STDOUT)
        rows.append({'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'log':str(log.relative_to(P)),'sha256':sha(log)})
        receipt={'source':before,'target':str(target),'rows':rows}
        (folder/'manifest.json').write_text(json.dumps(receipt,indent=2)+'\n');print(f'quality {attempt}/{i}: exit{r.returncode}',flush=True)
        assert r.returncode==0,rows[-1]
        assert census()==before,'source changed during quality gates'
    (P/'quality.json').write_text(json.dumps(receipt,indent=2)+'\n');print('PASS seven serial affected-owner gates')
