#!/usr/bin/env python3
"""Run the actual verifier with source custody before owned cleanup."""
import datetime,hashlib,importlib.util,json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;FORMAL=ROOT/'formal';REPO=ROOT.parents[3]
spec=importlib.util.spec_from_file_location('custody',FORMAL/'check.py');custody=importlib.util.module_from_spec(spec);spec.loader.exec_module(custody)
def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()
argv=[sys.executable,'-B',str(FORMAL/'verify.py'),'--precleanup']
record={'argv':argv,'cwd':str(REPO),'verifier_sha256':hashlib.sha256((FORMAL/'verify.py').read_bytes()).hexdigest(),'driver_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'source_before':custody.sources(),'started_utc':now()}
r=subprocess.run(argv,cwd=REPO,capture_output=True,text=True)
record.update(exit_code=r.returncode,finished_utc=now(),source_after=custody.sources(),stdout=r.stdout,stderr=r.stderr)
record['status']='pass' if r.returncode==0 and record['source_before']==record['source_after'] else 'failed'
with (ROOT/'precleanup.json').open('x') as f:f.write(json.dumps(record,indent=2)+'\n')
print(r.stdout+r.stderr,end='');assert record['status']=='pass'
