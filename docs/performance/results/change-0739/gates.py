"""Run the applicable documentation/coverage contract checks."""
import json,os,subprocess,time
from build import P,ROOT,sha
if __name__=='__main__':
    commands=[
      ['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--mode','structural'],
      ['python3','-B','tools/check_report_claim_classification.py'],
      ['python3','-B','tools/validate_crud_coverage_index.py'],
    ]
    rows=[]
    for i,cmd in enumerate(commands):
        log=P/f'gate-{i}.log';start=time.time()
        with log.open('w') as out:r=subprocess.run(cmd,cwd=ROOT,stdout=out,stderr=subprocess.STDOUT,env=os.environ|{'PYTHONDONTWRITEBYTECODE':'1'})
        rows.append({'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'log':log.name,'log_sha256':sha(log)})
        (P/'gates.json').write_text(json.dumps(rows,indent=2)+'\n')
        assert r.returncode==0,rows[-1]
    print('PASS three documentation/coverage structural gates; no new timing-contract promotion')
