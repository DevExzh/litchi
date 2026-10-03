#!/usr/bin/env python3
"""Single-scan symbol-quality preflight, separate from the formal population."""
from pathlib import Path
import datetime, importlib.util, json, os, subprocess, sys
P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('preflight_source_0894', P.parent/'change-0894/experiment.py')
e = importlib.util.module_from_spec(spec); sys.modules[spec.name] = e; spec.loader.exec_module(e)

def main():
    out = P/'preflight'
    assert not out.exists()
    plan = e.load_plan(P.parent/'change-0894/plan.json')
    binary = e._binary(plan, 'after', 'native')
    out.mkdir()
    argv = ['/usr/bin/taskset','-c','12','/usr/bin/heaptrack','--record-only','-o',str(out/'heaptrack'),binary['path'],'--case',e.PRIMARY_CASE,'--xlsx-shape','medium','--warmup','0','--samples','1','--filesystem-cache','warm','--json',str(out/'report.json'),'--producer-evidence',str(out/'producer.json')]
    env = dict(e.CHILD_ENV)
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    command = {'argv':argv,'cwd':str(P),'environment':env,'binary':binary,'started_utc':start,'scope':'whole-process one-scan symbol-quality preflight; excluded from formal population','helper':e.descriptor(Path(__file__))}
    (out/'command.json').write_text(json.dumps(command,indent=2)+'\n')
    with (out/'stdout.txt').open('xb') as stdout, (out/'stderr.txt').open('xb') as stderr:
        result = subprocess.run(argv,cwd=P,env=env,stdout=stdout,stderr=stderr)
    receipt = {**command,'exit_code':result.returncode,'finished_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'status':'fail','performance_claim':None}
    if result.returncode == 0:
        e._validate_child(json.loads((out/'report.json').read_text()),json.loads((out/'producer.json').read_text()),e.PRIMARY_CASE,'native',0,1,e._load_expected_corpora()[e.PRIMARY_CASE],binary)
        receipt['status']='pass'
    receipt['artifacts'] = {x.name:e.descriptor(x) for x in out.iterdir() if x.is_file()}
    (out/'receipt.json').write_text(json.dumps(receipt,indent=2)+'\n')
    assert receipt['status']=='pass'
    print('Single-scan preflight PASS; strict producer and per-worksheet guards valid.')
if __name__ == '__main__': main()
