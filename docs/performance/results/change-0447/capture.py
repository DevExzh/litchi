#!/usr/bin/env python3
"""Capture one matched pacing lane with exact source, binary and oracle custody."""
import argparse,datetime,hashlib,importlib.util,json,os,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3]
def sha(p):
    with p.open('rb') as f:return hashlib.file_digest(f,'sha256').hexdigest()
def load(p):return json.loads(p.read_text())
def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()
def command(binary,lane,report,samples,warmup,revision):
    cmd=[binary,'provider-lifecycle','--corpus',lane['corpus'],'--provider','range','--max-range','65536','--delay-us','200','--samples',str(samples),'--warmup',str(warmup),'--source-revision',revision,'--output',str(report)]
    if lane['paced']:cmd+=['--transfer-bytes-per-second','26214400']
    return cmd
def main():
    p=argparse.ArgumentParser();p.add_argument('--lane',type=int,required=True);p.add_argument('--pilot',action='store_true');p.add_argument('--profile',choices=['stat','record']);a=p.parse_args();assert not (a.pilot and a.profile)
    build=load(ROOT/'build.json');protocol=load(ROOT/'protocol.json') if not a.pilot else None
    lane=protocol['order'][a.lane] if protocol else {'corpus':('plain','media-rich')[a.lane//2],'paced':bool(a.lane%2),'repeat':'pilot'}
    directory=(ROOT/'profiles'/str(a.lane)/a.profile) if a.profile else (ROOT/'pilots'/str(a.lane)/'control' if a.pilot else ROOT/'runs'/str(a.lane));directory.mkdir(parents=True,exist_ok=False)
    report=directory/'report.json';resource=directory/'resource.log';log=directory/'workload.log';oracle_log=directory/'oracle.log'
    binary=Path(build['binary']['path']);assert sha(binary)==build['binary']['sha256']
    spec=importlib.util.spec_from_file_location('custody',ROOT/'check.py');custody=importlib.util.module_from_spec(spec);spec.loader.exec_module(custody)
    before=custody.sources();assert before==build['source_manifest']
    argv=['taskset','-c','2','/usr/bin/time','-v','-o',str(resource)]+command(str(binary),lane,report,1 if a.pilot else 30,0 if a.pilot else 3,build['revision'])
    perf_artifacts=[]
    if a.profile:
        assert lane['corpus']=='media-rich'
        resource_prefix=argv[:7];workload=argv[7:]
        if a.profile=='stat':
            output=directory/'perf-stat.txt';perf_artifacts=[('perf_stat',output)]
            perf=['perf','stat','--no-big-num','-x,','-e','cycles:u,instructions:u,branches:u,branch-misses:u,L1-dcache-load-misses:u','-o',str(output),'--']
        else:
            output=directory/'perf.data';perf_artifacts=[('perf_data',output)]
            perf=['perf','record','--no-buildid-cache','-o',str(output),'-F','999','-e','cycles:u','--call-graph','fp,127','--']
        argv=resource_prefix+perf+workload
    row={'change':447,'status':'running','lane':lane,'pilot':a.pilot,'profile':a.profile,'argv':argv,'source_before':before,'binary':build['binary'],'build_sha256':sha(ROOT/'build.json'),'protocol_sha256':None if a.pilot else sha(ROOT/'protocol.json'),'capture_sha256':sha(Path(__file__)),'oracle_sha256':sha(ROOT/'verify-report.py'),'started_utc':now()}
    target=directory/'receipt.json';target.write_text(json.dumps(row,indent=2)+'\n')
    try:
        with log.open('xb') as f:result=subprocess.run(argv,cwd=REPO,stdout=f,stderr=subprocess.STDOUT)
        row['exit_code']=result.returncode;assert result.returncode==0
        with oracle_log.open('xb') as f:result=subprocess.run([sys.executable,'-B',str(ROOT/'verify-report.py'),str(report)],stdout=f,stderr=subprocess.STDOUT)
        row['oracle_exit_code']=result.returncode;assert result.returncode==0
        if a.profile=='record':
            for key,extra in [('perf_report',['report','--stdio','--no-inline','-g','none']),('perf_script',['script','--no-inline'])]:
                path=directory/(key.replace('_','-')+'.txt');perf_artifacts.append((key,path))
                with path.open('xb') as f:subprocess.run(['perf',*extra,'-i',str(output)],stdout=f,stderr=subprocess.STDOUT,check=True)
        row['status']='pass'
    finally:
        row['source_after']=custody.sources();row['source_unchanged']=row['source_after']==before
        if not row['source_unchanged'] or row['status']=='running':row['status']='failed'
        row['finished_utc']=now();row['artifacts']={key:{'path':str(path.relative_to(ROOT)),'bytes':path.stat().st_size,'sha256':sha(path)} for key,path in [('report',report),('resource',resource),('workload',log),('oracle',oracle_log)]+perf_artifacts if path.exists()}
        target.write_text(json.dumps(row,indent=2)+'\n')
    assert row['status']=='pass';print(json.dumps({'status':'pass','lane':lane,'pilot':a.pilot}))
if __name__=='__main__':main()
