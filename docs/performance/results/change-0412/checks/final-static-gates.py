import datetime,json,os,subprocess,time
from pathlib import Path
root=Path('/home/zhuhe/code/litchi');out=Path('/tmp/litchi-goal-0412-checks');out.mkdir(exist_ok=True)
p=out/'final-static-commands.json';records=[]
env=os.environ.copy();env['RUSTUP_TOOLCHAIN']='1.98.1'
checks=[('rustfmt',['rustfmt','+1.98.1','--check','--edition','2024','--config','skip_children=true','tools/perf-baseline/src/lib.rs']),('boundaries',['python3','tools/check_crate_boundaries.py']),('coverage',['python3','tools/validate_crud_coverage_index.py']),('claims',['python3','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--evidence-root','.','--mode','strict']),('classification',['python3','tools/check_report_claim_classification.py','--repo-root','.'])]
for label,argv in checks:
 stamp=datetime.datetime.now(datetime.timezone.utc).isoformat();start=time.monotonic()
 with (out/('final-'+label+'.log')).open('wb') as log:r=subprocess.run(argv,cwd=root,env=env,stdout=log,stderr=subprocess.STDOUT)
 records.append(dict(label=label,argv=argv,cwd=str(root),started_utc=stamp,wall_seconds=time.monotonic()-start,exit_code=r.returncode));p.write_text(json.dumps(records,indent=2)+'\n');print(label,r.returncode,flush=True);r.check_returncode()
