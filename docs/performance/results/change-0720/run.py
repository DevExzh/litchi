#!/usr/bin/env python3
"""Serial diagnostic commands with source and artifact custody; no timings claimed."""
import json, os, subprocess, sys, time
from custody import P, ROOT, TARGET, census, sha
stage=sys.argv[1]
manifest=P/'oracle/Cargo.toml'
binary=TARGET/'debug/docx-active-offset-oracle-0720'
commands={
 'baseline-build':['cargo','build','--locked','--manifest-path',str(manifest)],
 'trace-build':['cargo','build','--locked','--manifest-path',str(manifest)],
 'baseline':[str(binary),'--output',str(P/'baseline.json')],
 'trace':[str(binary),'--output',str(P/'trace.json')],
 'trace-repeat':[str(binary),'--output',str(P/'trace-repeat.json')],
}
command=commands[stage]
receipt=P/(stage+'-receipt.json')
assert not receipt.exists()
before=census()
(P/(stage+'-source.json')).write_text(json.dumps(before,indent=2)+'\n')
env=os.environ.copy();env['CARGO_TARGET_DIR']=str(TARGET);env['CARGO_BUILD_JOBS']='4'
start=time.time()
with (P/(stage+'.stdout')).open('wb') as out,(P/(stage+'.stderr')).open('wb') as err:
 result=subprocess.run(command,cwd=ROOT,env=env,stdout=out,stderr=err)
after=census()
assert before==after
record={'stage':stage,'argv':command,'cwd':str(ROOT),'exit_code':result.returncode,'elapsed_seconds':time.time()-start,'source_sha256':sha(P/(stage+'-source.json')),'binary_sha256':sha(binary) if binary.exists() else None,'oracle_source_sha256':sha(P/'oracle/src/main.rs'),'stdout_sha256':sha(P/(stage+'.stdout')),'stderr_sha256':sha(P/(stage+'.stderr')),'target':str(TARGET),'cargo_build_jobs':4}
if (P/(stage+'.json')).exists():record['report_sha256']=sha(P/(stage+'.json'))
receipt.write_text(json.dumps(record,indent=2)+'\n')
print(json.dumps(record));sys.exit(result.returncode)
