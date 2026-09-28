"""Reuse exact-source sealed baseline quality before a fresh candidate trial."""
from pathlib import Path
import custody as c

P=c.P;OLD=P.parent/'change-0810';OUT=P/'quality-reuse'
assert not OUT.exists() and not (P/'quality-reuse.json').exists()
source=c.source();prior=c.read(OLD/'build-after/source.json')
assert source['files']==prior['files'] and len(source['files'])==9196
assert source['revision']==c.read(P/'origin.json')['base']
seal=c.read(OLD/'seal.json')['files']
def sealed(path):
    row=c.artifact(path)
    assert seal[str(Path(path).relative_to(c.ROOT))]==row['sha256']
    return row
summary=c.read(OLD/'quality-summary.json')['production_after']
assert summary['schema']=='litchi.performance.0810.quality-after.v1'
assert len(summary['gates'])==6
assert {k:summary['tests'][k] for k in ('passed','failed','ignored','suites')}=={'passed':1241,'failed':0,'ignored':3,'suites':85}
assert all(g['exit_code']==0 and g['status']=='pass' for g in summary['gates'])
root_inputs=c.assert_root_inputs();assert summary['root_inputs']==root_inputs
checks=sealed(summary['checks']['path'])
for gate in summary['gates']:assert sealed(gate['log']['path'])==gate['log']
architecture=c.read(P/'architecture-inputs.json')
assert architecture==c.read(P.parent/'change-0812/architecture-inputs.json')
for name,digest in architecture.items():assert c.sha(c.ROOT/name)==digest
unrelated=c.read(P.parent/'change-0811/build/frozen-inputs.json')['unrelated']
for name,digest in unrelated.items():assert c.sha(c.ROOT/name)==digest
probe={str(path.relative_to(P/'probe-src')):c.sha(path) for path in (P/'probe-src').rglob('*') if path.is_file()}
assert probe==c.read(P.parent/'change-0811/build/probe.json')
OUT.mkdir()
c.write(OUT/'reuse-inputs.json',{'current_source':source,'sealed_0810_source':sealed(OLD/'build-after/source.json'),
    'sealed_0810_summary':sealed(OLD/'quality-summary.json'),'sealed_0810_source_files_equal':True,
    'root_inputs':root_inputs,'architecture':architecture,'unrelated':unrelated,'probe':probe})
c.write(P/'quality-reuse.json',{'schema':'litchi.performance.0813.quality-reuse.v1',
    'mode':'reuse-sealed-0810-production-gates','reference':sealed(OLD/'quality-summary.json'),
    'source':sealed(OLD/'build-after/source.json'),'source_files_equal':True,'gate_count':6,
    'gates':summary['gates'],'tests':summary['tests'],'inputs':c.artifact(OUT/'reuse-inputs.json'),
    'root_inputs':root_inputs,'architecture':architecture,'unrelated':unrelated,'cargo_executed':False,
    'reason':'0813 baseline equals all 9196 source hashes of the sealed 0810 after source'})
print('0813 baseline quality reuse PASS')
