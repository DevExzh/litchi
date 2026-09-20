#!/usr/bin/env python3
"""Validate restored production and retained diagnostic/documentation."""
import json, os, subprocess, time
from custody import P, ROOT, TARGET, census, sha
baseline=json.loads((P/'baseline-source.json').read_text())
assert census()==baseline
manifest=str(P/'oracle/Cargo.toml')
commands=[('fmt',['cargo','fmt','--manifest-path',manifest,'--','--check']),('clippy',['cargo','clippy','--locked','--manifest-path',manifest,'--all-targets','--','-D','warnings'])]
commands += [(r['name'],r['command']) for r in json.loads((P.parent/'change-0717/evidence/results.json').read_text())]
results=[]
for name,argv in commands:
 log=P/('quality-'+name+'.log'); assert not log.exists()
 start=time.monotonic()
 with log.open('x') as out:
  result=subprocess.run(argv,cwd=ROOT,env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='4'),stdout=out,stderr=subprocess.STDOUT)
 row={'name':name,'argv':argv,'exit_code':result.returncode,'seconds':time.monotonic()-start,'log_sha256':sha(log)}
 results.append(row);(P/'quality.json').write_text(json.dumps(results,indent=2)+'\n')
 print(name,result.returncode,flush=True)
 assert result.returncode==0
assert census()==baseline
