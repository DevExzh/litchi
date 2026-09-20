#!/usr/bin/env python3
"""Serial DOCX quality and repository evidence gates, with source custody."""
import json,os,subprocess,sys,time
from custody import P,ROOT,TARGET,census,sha
mode=sys.argv[1];assert mode in ['docx','harness','evidence']
source=census();manifest=P/('quality-'+mode+'-source.json');assert not manifest.exists();manifest.write_text(json.dumps(source,indent=2)+'\n')
common=['--locked','--target-dir',str(TARGET),'-j','2']
if mode=='docx':
 jobs=[('fmt',['cargo','fmt','-p','litchi-docx','--','--check']),
 ('tests',['cargo','test','-p','litchi-docx',*common,'--all-features','--all-targets']),
 ('clippy',['cargo','clippy','-p','litchi-docx',*common,'--all-features','--all-targets','--','-D','warnings']),
 ('doctests',['cargo','test','-p','litchi-docx',*common,'--all-features','--doc']),
 ('rustdoc',['cargo','doc','-p','litchi-docx',*common,'--all-features','--no-deps'])]
elif mode=='harness':
 jobs=[('tests',['cargo','test','--release','--manifest-path','tools/perf-baseline/Cargo.toml',*common,'--lib'])]
else:jobs=[(r['name'],r['command']) for r in json.loads((P.parent/'change-0717/evidence/results.json').read_text())]
rows=[]
for name,cmd in jobs:
 log=P/f'quality-{mode}-{name}.log';assert not log.exists();started=time.monotonic()
 with log.open('x') as out:
  result=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,**({'RUSTDOCFLAGS':'-D warnings'} if name=='rustdoc' else {})),stdout=out,stderr=subprocess.STDOUT)
 assert census()==source
 rows.append({'name':name,'command':cmd,'exit_code':result.returncode,'seconds':time.monotonic()-started,'log':log.name,'log_sha256':sha(log),'source_manifest_sha256':sha(manifest)})
 (P/f'quality-{mode}.json').write_text(json.dumps(rows,indent=2)+'\n');print(mode,name,result.returncode,flush=True)
 assert result.returncode==0
