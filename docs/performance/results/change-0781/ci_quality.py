"""Root offline gates for the workflow matrix repair, retaining every attempt."""
import subprocess,time
import custody as c
attempt=0
while (c.P/f'ci-quality-{attempt}').exists():attempt+=1
out=c.P/f'ci-quality-{attempt}';out.mkdir()
names=['.github/workflows/perf-baseline.yml','docs/performance/crud-coverage-index-v2.json','docs/performance/results/perf-regression-default-manifest-v1.json','docs/performance/results/perf-corpus-manifest-v2.json']+[str(f.relative_to(c.ROOT)) for f in sorted((c.ROOT/'tools').glob('*.py'))]
frozen={n:c.sha(c.ROOT/n) for n in names};c.write(out/'source.json',frozen)
commands=[
 ['python3','-B','-m','unittest','tools.test_validate_perf_default_matrix','tools.test_perf_workflow_policy','tools.test_crud_coverage_index'],
 ['python3','-B','tools/validate_perf_default_matrix.py','--manifest','docs/performance/results/perf-regression-default-manifest-v1.json','--report',str(c.P/'ci-capture/smoke.json'),'--mode','smoke','--samples','2','--shape','tiny','--payload','compressible'],
 ['python3','-B','tools/validate_perf_default_matrix.py','--manifest','docs/performance/results/perf-regression-default-manifest-v1.json','--report',str(c.P/'ci-capture/full.json'),'--mode','full','--samples','15'],
 ['python3','-B','tools/validate_crud_coverage_index.py','--index','docs/performance/crud-coverage-index-v2.json'],
 ['python3','-B','tools/validate_perf_corpus_binding.py','--report',str(c.P/'ci-capture/full.json'),'--catalog',str(c.P/'ci-capture/full.corpus-manifest-v2.json')],
 ['python3','-B','tools/validate_crud_coverage_index.py','--index','docs/performance/crud-coverage-index-v2.json','--catalog',str(c.P/'ci-capture/full.corpus-manifest-v2.json'),'--selector-source','tools/perf-baseline/src/lib.rs','--checklist','docs/CRUD_Scenario_Checklist.md','--repo-root','.','--report',str(c.P/'ci-capture/full.json')],
]
rows=[]
for i,cmd in enumerate(commands):
 log=out/f'{i:02}.log';start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
 rows.append({'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'log':c.artifact(log)});c.write(out/'checks.json',rows)
 print('CI gate',i+1,'exit',r.returncode,flush=True)
 assert r.returncode==0,log
 assert {n:c.sha(c.ROOT/n) for n in names}==frozen
c.write(c.P/'ci-quality.json',{'rows':rows,'source':c.artifact(out/'source.json')})
