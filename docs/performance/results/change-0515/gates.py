"""Checks appropriate to an unchanged-production profiling batch."""
import subprocess,datetime,time,json
from run import HERE,REPO,sha,sources,write
commands=[('fmt',['cargo','fmt','--all','--','--check']),('boundaries',['python3','-B','tools/check_crate_boundaries.py']),('claims',['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--evidence-root','.','--mode','strict'])]
for name,cmd in commands:
    manifest=sources();assert manifest==json.loads((HERE/'source-manifest.json').read_text())
    started=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
    with (HERE/(name+'.log')).open('x') as out:r=subprocess.run(cmd,cwd=REPO,stdout=out,stderr=subprocess.STDOUT)
    unchanged=sources()==manifest
    write(HERE/(name+'-receipt.json'),{'command':cmd,'started_utc':started,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':unchanged,'source_manifest_sha256':sha(HERE/'source-manifest.json'),'log_sha256':sha(HERE/(name+'.log'))})
    print(name,r.returncode,flush=True);assert r.returncode==0 and unchanged
