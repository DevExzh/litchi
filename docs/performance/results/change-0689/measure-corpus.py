#!/usr/bin/env python3
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];phase=sys.argv[1]
assert phase in ['baseline','candidate']
suffix='before' if phase=='baseline' else 'after'
source=ROOT
binary=Path('/home/zhuhe/code/litchi-target-0689-'+suffix)/'release/xls-index-probe-0684'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
out=P/'corpus'/phase;out.mkdir(parents=True,exist_ok=True)
manifest=dict(binary_sha256=sha(binary),source_sha256={str(f.relative_to(source)):sha(f) for owner in ['litchi-cfb','litchi-xls'] for f in sorted((source/'crates'/owner).rglob('*.rs'))},probe_sha256={str(f.relative_to(ROOT)):sha(f) for f in sorted((P.parent/'change-0684/probe').rglob('*')) if f.is_file()},corpus_manifest_sha256=sha(P/'corpus-manifest.json'),commands=[])
for mode in ['owned','file']:
 cmd=['taskset','-c','12',str(binary),'corpus','--root','test-data','--mode',mode,'--sample-coordinates','16','--max-queries','512']
 r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
 (out/(mode+'.json')).write_text(r.stdout)
 j=json.loads(r.stdout);assert j['query_mismatches']==0
 if phase=='candidate':assert j==json.loads((P/'corpus/baseline'/(mode+'.json')).read_text())
 manifest['commands'].append(dict(command=cmd,exit_code=r.returncode))
 print(phase,mode,j['files_seen'],flush=True)
for c in json.loads((P/'cases.json').read_text()):
 if not c['case'].startswith('synthetic-'):continue
 for mode in ['owned','file']:
  cmd=['taskset','-c','12',str(binary),'route','--input',c['path'],'--route','visit','--mode',mode,'--worksheet','0','--row','0','--column','0','--samples','1','--warmups','1']
  r=subprocess.run(cmd,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
  j=json.loads(r.stdout);name=c['case']+'-'+mode+'.json';(out/name).write_text(json.dumps(j,separators=(',',':'))+'\n')
  expected=int(c['case'].split('-')[1])+1
  assert j['records'][0]['visit']['actual_callbacks']==expected
  assert j['records'][0]['visit']['oracle_callbacks']==expected
  manifest['commands'].append(dict(command=cmd,exit_code=r.returncode))
  print(phase,c['case'],mode,expected,flush=True)
manifest['generated_fixtures']={c['path']:sha(ROOT/c['path']) for c in json.loads((P/'cases.json').read_text()) if c['case'].startswith('synthetic-')}
manifest['raw_sha256']={f.name:sha(f) for f in out.iterdir() if f.is_file() and f.name!='manifest.json'}
(out/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
