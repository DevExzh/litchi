from pathlib import Path
import gzip, hashlib, json, os, subprocess, sys
os.environ["RUSTUP_TOOLCHAIN"]="1.98.1"
sys.path.insert(0,str(Path.cwd()/'tools'))
import perf_compare
out=Path('docs/performance/results/change-0405');out.mkdir(parents=True,exist_ok=True)
binary=Path('tools/perf-baseline/target/release/litchi-perf-baseline').resolve()
revision=subprocess.check_output(["git","rev-parse","HEAD"],text=True).strip()
subprocess.run(["git","diff","--quiet","HEAD"],check=True)
checks=[]
for observed in [False,True]:
 name='observed' if observed else 'normal'
 report_path=out/(name+'.json')
 command=[str(binary),'--case','opc_source_cache_control_contention,opc_source_cache_managed_contention','--workers','1,2','--warmup','1','--samples','3','--json',str(report_path)]
 if observed: command+=['--opc-cache-lock-diagnostics']
 result=subprocess.run(command,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
 (out/(name+'.log.gz')).write_bytes(gzip.compress(result.stdout,mtime=0))
 if result.returncode: raise RuntimeError(result.stdout.decode())
 report=json.loads(report_path.read_text());perf_compare.validate_parallel_metrics(report)
 assert len(report['results'])==24
 assert report['configuration']['opc_cache_lock_diagnostics'] is observed
 assert report['environment']['git_revision']==revision
 assert '1.98.1' in report['environment']['rustc_version']
 for row in report['results']:
  diagnostic=row['source']['opc_cache'].get('lock_diagnostics')
  assert (diagnostic is not None) is observed
 checks.append(dict(command=command,exit_code=result.returncode,report=report_path.name,result_rows=len(report['results']),parallel_metrics_validation='passed'))
command=[str(binary),'--case','opc_open','--warmup','0','--samples','1','--opc-cache-lock-diagnostics']
result=subprocess.run(command,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
assert result.returncode!=0 and b'requires an OPC source-cache contention case' in result.stdout
(out/'invalid-selector.log.gz').write_bytes(gzip.compress(result.stdout,mtime=0))
checks.append(dict(command=command,exit_code=result.returncode,expected_rejection=True))
manifest=dict(schema=1,kind='diagnostic smoke only',performance_claim='none',claim_authorized=False,revision=subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip(),binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest(),toolchain=subprocess.check_output(['rustc','+1.98.1','-vV'],text=True),checks=checks,source_sha256={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in map(Path,['tools/perf-baseline/Cargo.toml','tools/perf-baseline/Cargo.lock','tools/perf-baseline/src/lib.rs','tools/perf-baseline/src/parallel_metrics.rs','tools/perf_compare.py','tools/test_perf_compare.py'])})
manifest['git_status']=subprocess.check_output(['git','status','--porcelain=v1'],text=True)
manifest['source_state']='committed tracked source; untracked goal and pending evidence are retained'
manifest['clean_abba_claim_eligible']=False
manifest['capture_driver']='capture-smoke.py'
(out/'capture-smoke.py').write_text(Path(__file__).read_text())
(out/'validation.json').write_text(json.dumps(manifest,indent=2)+'\n')
print([(x.get('report'),x.get('result_rows'),x['exit_code']) for x in checks])
