import hashlib,json,subprocess,os
from pathlib import Path
base=Path(__file__).resolve().parent
packet=base/'verified-extraction/litchi-performance-0947'
rows=[]
def run(name,argv,env=None,expected=0):
 r=subprocess.run(argv,cwd=packet,env=env,text=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
 for suffix,value in [('stdout',r.stdout),('stderr',r.stderr)]:
  with (base/f'{name}.{suffix}').open('x') as f:f.write(value)
 row={'name':name,'argv':argv,'cwd':str(packet),'returncode':r.returncode,'expected_returncode':expected}
 rows.append(row)
 with (base/f'{name}.json').open('x') as f:json.dump(row,f,indent=2);f.write('\n')
 assert r.returncode==expected,(name,r.stderr)
def digest(p):return hashlib.sha256(p.read_bytes()).hexdigest()
env=dict(os.environ,CARGO_TARGET_DIR=str(base/'rebuild-target'))
run('moved-build',['cargo','+1.95.0','build','--manifest-path','probe-v2/Cargo.toml','--locked','--offline'],env)
binary=base/'rebuild-target/debug/litchi-proof-diagnostic'
proof=base/'replayed-proof.tsv';diagnostic=base/'replayed-diagnostic.json';model=base/'replayed-model.json'
run('moved-proof',[str(binary),'inputs/artifacts',str(proof)])
assert proof.read_bytes()==(packet/'results/proof-portable-final.tsv').read_bytes()
run('moved-analysis',['python3','-B','analysis/analyze.py','--artifacts','inputs/artifacts','--manifest','inputs/artifacts/manifest.json','--probe',str(proof),'--output',str(diagnostic)])
assert diagnostic.read_bytes()==(packet/'results/diagnostic-portable-final.json').read_bytes()
run('moved-model',['python3','-B','identity/coverage_model.py','--spans','review/changed-cell-span-review-v4.json','--proof',str(proof),'--output',str(model)])
assert model.read_bytes()==(packet/'identity/coverage-model-v1.json').read_bytes()
run('probe-overwrite-control',[str(binary),'inputs/artifacts',str(proof)],expected=1)
assert proof.read_bytes()==(packet/'results/proof-portable-final.tsv').read_bytes()
manifest=json.loads((base/'inventory-v1.json').read_text())['files']
actual=sorted(p.relative_to(packet).as_posix() for p in packet.rglob('*') if p.is_file())
assert actual==sorted(r['path'] for r in manifest)
for r in manifest:
 f=packet/r['path'];assert f.stat().st_size==r['bytes'] and digest(f)==r['sha256']
result={'status':'pass','scope':'Actual relocated locked offline probe rebuild and byte-identical TSV, analyzer, and model replay; overwrite refusal preserves output. All extracted files remain unchanged.','commands':rows,'rebuilt_binary':{'bytes':binary.stat().st_size,'sha256':digest(binary)},'outputs':{p.name:{'bytes':p.stat().st_size,'sha256':digest(p)} for p in [proof,diagnostic,model]},'replay_script_sha256':digest(Path(__file__))}
with (base/'root-replay-v1.json').open('x') as f:json.dump(result,f,indent=2);f.write('\n')
print(json.dumps({'status':'pass','outputs':result['outputs']}))
