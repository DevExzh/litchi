"""Seal the reviewed, restored, cleaned documentation-only batch before commit."""
import pathlib,json,hashlib,subprocess,ast
p=pathlib.Path(__file__).resolve().parent;root=p.parents[3]
def d(f):
 raw=f.read_bytes();return {'path':str(f.relative_to(root)),'bytes':len(raw),'sha256':hashlib.sha256(raw).hexdigest()}
r=json.loads((p/'result-review.json').read_text());assert r['status']=='pass'
records=r['allocator']['records_detail'];vectors=sum(len(x.get('phase_metrics',{})) for x in records);samples=sum(len(v) for x in records for v in x.get('phase_metrics',{}).values())
assert vectors==r['allocator']['phase_vectors_checked']==60;assert samples==r['allocator']['sample_metric_records_checked']==180
assert json.loads((p/'post-cleanup-replay.json').read_text())['status']=='pass'
assert json.loads((p/'cleanup.json').read_text())['status']=='pass'
assert json.loads((p/'decision.json').read_text())['status']=='rejected'
for name in ['/tmp/litchi-0902-src','/home/zhuhe/code/litchi/target/litchi-0902-build']:assert not pathlib.Path(name).exists()
manifest=json.loads((p.parent/'change-0894'/'source.json').read_text())['files']
for name,expected in manifest.items():
 actual=d(root/name);assert actual['bytes']==expected['bytes'] and actual['sha256']==expected['sha256'],name
for expected in json.loads((p/'unrelated-preservation.json').read_text()):assert d(root/expected['path'])==expected
for f in p.rglob('*.py'):ast.parse(f.read_text(),filename=str(f))
files=subprocess.check_output(['git','diff','--cached','--name-only'],cwd=root,text=True).splitlines()
allowed={'docs/performance/'+s for s in ['BASELINE.md','HOTSPOTS.md','REPORT.md','GOAL_AUDIT.md','0902-xlsx-upfront-stopped-audit-candidate.md']}
assert files and all(f in allowed or f.startswith('docs/performance/results/change-0902/') for f in files)
assert not subprocess.check_output(['git','diff','--name-only','--',*files],cwd=root)
exceptions=json.loads((p/'staging-whitespace-check.json').read_text())['immutable_tool_output_exceptions']
subprocess.run(['git','diff','--cached','--check','--','.',*[f':(exclude){f}' for f in exceptions]],cwd=root,check=True)
head=subprocess.check_output(['git','rev-parse','HEAD'],cwd=root,text=True).strip();assert head=='23bf87661ed9fe17c15d8cf4ab8f6e478138347d'
result={'status':'pass','parent_head':head,'decision':'rejected; production restored','baseline_source_files_verified':len(manifest),'worktrees_created':0,'owned_roots_absent':True,'unrelated_files_preserved':True,'independent_review':d(p/'result-review.json'),'bindings':[d(root/f) for f in files],'bound_files_before_closure':len(files),'closure_excluded_from_own_bindings':True,'ignored_binary_artifacts':'retained by plan/capture/replay SHA256 custody, intentionally outside Git'}
with (p/'closure.json').open('x') as out:json.dump(result,out,indent=2,sort_keys=True);out.write('\n')
print('sealed',len(files),'files; all10771 baseline source files restored')
