from pathlib import Path
import subprocess,json,os,time
root=Path('/home/zhuhe/code/litchi');out=Path('/tmp/litchi-goal-0413-checks');env=os.environ.copy();env.update(RUSTUP_TOOLCHAIN='1.98.1',PYTHONDONTWRITEBYTECODE='1')
commands=[['python3','tools/check_crate_boundaries.py'],['python3','tools/validate_crud_coverage_index.py'],['python3','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--evidence-root','.','--mode','strict'],['python3','tools/check_report_claim_classification.py','--repo-root','.']]
records=[]
for argv in commands:
 label=Path(argv[1]).stem;start=time.monotonic()
 with (out/(label+'.log')).open('wb') as f:r=subprocess.run(argv,cwd=root,env=env,stdout=f,stderr=subprocess.STDOUT)
 records.append(dict(argv=argv,exit_code=r.returncode,wall_seconds=time.monotonic()-start));(out/'final-gates.json').write_text(json.dumps(records,indent=2)+'\n');print(label,r.returncode,flush=True);r.check_returncode()
