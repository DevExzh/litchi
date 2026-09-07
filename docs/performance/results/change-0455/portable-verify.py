#!/usr/bin/env python3
"""Replay a separate evidence copy after owned build/output artifacts are gone."""
import datetime,hashlib,json,os,shutil,subprocess,sys,tempfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent
assert not Path('/tmp/litchi-goal-0455').exists()
started=datetime.datetime.now(datetime.timezone.utc).isoformat()
with tempfile.TemporaryDirectory(prefix='litchi-goal-0455-portable-') as directory:
    parent=Path(directory);target=parent/'docs/performance/results/change-0455';shutil.copytree(ROOT,target)
    argv=[sys.executable,'-B',str(target/'formal/verify.py')]
    result=subprocess.run(argv,cwd=parent,env=os.environ|{'PYTHONPATH':'','PYTHONDONTWRITEBYTECODE':'1'},capture_output=True,text=True)
    record={'status':'pass' if result.returncode==0 else 'failed','argv':argv,'cwd':str(parent),'exit_code':result.returncode,'started_utc':started,'finished_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'verifier_sha256':hashlib.sha256((target/'formal/verify.py').read_bytes()).hexdigest(),'driver_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'mode':'separate bundle copy without owned binaries, fuzz build, native outputs or raw perf data','stdout':result.stdout,'stderr':result.stderr}
    if result.returncode==0:assert json.loads(result.stdout)['status']=='pass'
record['temporary_directory_absent']=not parent.exists()
with (ROOT/'portable-verification.json').open('x') as f:f.write(json.dumps(record,indent=2)+'\n')
print(result.stdout+result.stderr,end='');sys.exit(result.returncode)
