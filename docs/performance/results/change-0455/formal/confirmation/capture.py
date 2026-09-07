#!/usr/bin/env python3
"""Capture one immutable lane with explicit executable and source identities."""
import argparse, datetime, hashlib, importlib.util, json, subprocess, sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[5]
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def load(p): return json.loads(p.read_text())
def now(): return datetime.datetime.now(datetime.timezone.utc).isoformat()
def artifact(p): return {'path':str(p.relative_to(ROOT)), 'sha256':sha(p), 'bytes':p.stat().st_size}
def main():
    ap=argparse.ArgumentParser();ap.add_argument('lane',type=int);a=ap.parse_args()
    protocol=load(ROOT/'protocol.json');lane=protocol['lanes'][a.lane]
    for name,digest in protocol['bound_files'].items(): assert sha(ROOT/name)==digest,(name,'driver changed')
    buildpath=ROOT/(lane['build']+'-build.json');build=load(buildpath)
    binary=build['binaries'][lane['instrumentation']];bp=Path(binary['path'])
    assert bp.is_file() and sha(bp)==binary['sha256'] and bp.stat().st_size==binary['bytes']
    spec=importlib.util.spec_from_file_location('custody',ROOT/'check.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
    before=mod.sources();known=[load(ROOT/(name+'-build.json'))['source_manifest'] for name in ['baseline','candidate'] if (ROOT/(name+'-build.json')).exists()]
    assert before in known,'unbuilt current source epoch'
    directory=ROOT/'runs'/str(a.lane);directory.mkdir(parents=True,exist_ok=False)
    report=directory/'report.json';resource=directory/'resource.log'
    argv=['taskset','-c',str(protocol['cpu']),'/usr/bin/time','-v','-o',str(resource),str(bp),'provider-lifecycle','--corpus',lane['corpus'],'--provider',lane['provider'],'--samples',str(protocol['samples']),'--warmup',str(protocol['warmups']),'--source-revision',build['revision'],'--output',str(report)]
    if lane['provider']=='range':
        argv+=['--max-range','65536','--delay-us','200','--transfer-bytes-per-second','26214400','--transfer-delay-policy','separate-sleeps']
    receipt=directory/'receipt.json';record={'status':'running','lane':a.lane,'lane_definition':lane,'argv':argv,'cwd':str(REPO),'started_utc':now(),'binary':binary,'build_manifest':artifact(buildpath),'source_before':before,'protocol_sha256':sha(ROOT/'protocol.json'),'capture_sha256':sha(Path(__file__))}
    receipt.write_text(json.dumps(record,indent=2)+'\n')
    try:
        with (directory/'workload.log').open('xb') as out:result=subprocess.run(argv,cwd=REPO,stdout=out,stderr=subprocess.STDOUT)
        record['exit_code']=result.returncode;assert result.returncode==0,'workload failed'
        oracle_argv=[sys.executable,'-B',str(ROOT/'verify-report.py'),str(report)];record['oracle_argv']=oracle_argv
        with (directory/'oracle.log').open('xb') as out:result=subprocess.run(oracle_argv,cwd=REPO,stdout=out,stderr=subprocess.STDOUT)
        record['oracle_exit_code']=result.returncode;assert result.returncode==0,'oracle failed'
        record['status']='pass'
    finally:
        record.update(finished_utc=now(),source_after=mod.sources())
        record['source_unchanged']=record['source_after']==before
        if record['status']=='running' or not record['source_unchanged']:record['status']='failed'
        record['artifacts']={p.name:artifact(p) for p in [report,resource,directory/'workload.log',directory/'oracle.log'] if p.exists()}
        receipt.write_text(json.dumps(record,indent=2)+'\n')
    assert record['status']=='pass';print(json.dumps({'status':'pass','lane':a.lane}),flush=True)
if __name__=='__main__':main()
