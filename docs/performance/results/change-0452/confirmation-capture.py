#!/usr/bin/env python3
"""Run one isolated baseline/candidate lifecycle process with bound binary custody."""
import argparse,datetime,hashlib,importlib.util,json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3]
def sha(p):
    with p.open('rb') as f:return hashlib.file_digest(f,'sha256').hexdigest()
def load(p):return json.loads(p.read_text())
def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()
def command(binary,lane,report,samples,warmup,revision):
    cmd=[binary,'provider-lifecycle','--corpus',lane['corpus'],'--provider',lane['provider'],'--samples',str(samples),'--warmup',str(warmup),'--source-revision',revision,'--output',str(report)]
    if lane['provider']=='range':cmd+=['--max-range','65536','--delay-us','200','--transfer-bytes-per-second','26214400','--transfer-delay-policy','separate-sleeps']
    return cmd
def main():
    p=argparse.ArgumentParser();p.add_argument('--lane',type=int,required=True);p.add_argument('--pilot',action='store_true');p.add_argument('--profile',choices=['stat','record']);a=p.parse_args();assert not(a.pilot and a.profile)
    protocol=load(ROOT/'confirmation-protocol.json');lane=protocol['order'][a.lane]
    build=load(ROOT/(lane['build']+'-build.json'))
    directory=ROOT/'confirmation'/str(a.lane)
    if a.profile:directory/=a.profile
    directory.mkdir(parents=True,exist_ok=False)
    report=directory/'report.json';resource=directory/'resource.log';log=directory/'workload.log';oracle_log=directory/'oracle.log'
    binary=Path(build['binary']['path']);assert sha(binary)==build['binary']['sha256']
    spec=importlib.util.spec_from_file_location('custody',ROOT/'check.py');c=importlib.util.module_from_spec(spec);spec.loader.exec_module(c)
    before=c.sources();assert before==load(ROOT/'candidate-build.json')['source_manifest']
    workload=command(str(binary),lane,report,1 if a.pilot else protocol['samples'],0 if a.pilot else protocol['warmups'],build['revision'])
    perf_artifacts=[]
    if a.profile:
        if a.profile=='stat':
            output=directory/'perf-stat.txt';perf_artifacts=[('perf_stat',output)]
            perf=['perf','stat','--no-big-num','-x,','-e','cycles:u,instructions:u,branches:u,branch-misses:u,L1-dcache-load-misses:u','-o',str(output),'--']
        else:
            output=directory/'perf.data';perf_artifacts=[('perf_data',output)]
            perf=['perf','record','--no-buildid-cache','-o',str(output),'-F','999','-e','cycles:u','--call-graph','fp,127','--']
        workload=perf+workload
    argv=['taskset','-c',str(protocol['cpu']),'/usr/bin/time','-v','-o',str(resource)]+workload
    row={'change':452,'status':'running','lane':lane,'pilot':a.pilot,'profile':a.profile,'argv':argv,'source_before':before,'binary':build['binary'],'build_sha256':sha(ROOT/(lane['build']+'-build.json')),'protocol_sha256':sha(ROOT/'confirmation-protocol.json'),'capture_sha256':sha(Path(__file__)),'oracle_sha256':sha(ROOT/'verify-report.py'),'started_utc':now()}
    target=directory/'receipt.json';target.write_text(json.dumps(row,indent=2)+'\n')
    try:
        with log.open('xb') as f:r=subprocess.run(argv,cwd=REPO,stdout=f,stderr=subprocess.STDOUT)
        row['exit_code']=r.returncode;assert r.returncode==0
        with oracle_log.open('xb') as f:r=subprocess.run([sys.executable,'-B',str(ROOT/'verify-report.py'),str(report)],stdout=f,stderr=subprocess.STDOUT)
        row['oracle_exit_code']=r.returncode;assert r.returncode==0
        if a.profile=='record':
            for key,extra in [('perf_report',['report','--stdio','--no-inline','-g','none']),('perf_script',['script','--no-inline'])]:
                path=directory/(key.replace('_','-')+'.txt');perf_artifacts.append((key,path))
                with path.open('xb') as f:subprocess.run(['perf',*extra,'-i',str(output)],stdout=f,stderr=subprocess.STDOUT,check=True)
        row['status']='pass'
    finally:
        row['source_after']=c.sources();row['source_unchanged']=row['source_after']==before
        if not row['source_unchanged'] or row['status']=='running':row['status']='failed'
        row['finished_utc']=now();row['artifacts']={key:{'path':str(path.relative_to(ROOT)),'bytes':path.stat().st_size,'sha256':sha(path)} for key,path in [('report',report),('resource',resource),('workload',log),('oracle',oracle_log)]+perf_artifacts if path.exists()}
        target.write_text(json.dumps(row,indent=2)+'\n')
    assert row['status']=='pass';print(json.dumps({'status':'pass','lane':lane,'pilot':a.pilot,'profile':a.profile}))
if __name__=='__main__':main()
