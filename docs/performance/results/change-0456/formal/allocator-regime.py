#!/usr/bin/env python3
"""Declared diagnostic: fixed mmap thresholds, ABBA each, no primary-lane substitution."""
import hashlib,importlib.util,json,os,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[4]
spec=importlib.util.spec_from_file_location('custody',ROOT/'check.py');custody=importlib.util.module_from_spec(spec);spec.loader.exec_module(custody)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
    protocol=json.loads((ROOT/'allocator-regime-protocol.json').read_text())
    assert sha(Path(__file__))==protocol['driver_sha256']
    inherited={k:v for k,v in os.environ.items() if k.startswith('MALLOC_') or k=='GLIBC_TUNABLES'}
    assert inherited==protocol['inherited_allocator_environment']=={}
    rows=[]
    for i,lane in enumerate(protocol['lanes']):
        build=json.loads((ROOT/(lane['build']+'-build.json')).read_text());binary=build['binaries']['normal'];assert sha(Path(binary['path']))==binary['sha256']
        directory=ROOT/'allocator-regime'/str(i);directory.mkdir(parents=True,exist_ok=False);report=directory/'report.json'
        argv=['taskset','-c','2','/usr/bin/time','-v','-o',str(directory/'resource.log'),'perf','stat','-x,','-e',protocol['events'],'-o',str(directory/'perf.csv'),binary['path'],'provider-lifecycle','--corpus','media-rich','--provider','bytes','--samples','30','--warmup','3','--source-revision',build['revision'],'--output',str(report)]
        before=custody.sources();env={'MALLOC_MMAP_THRESHOLD_':str(lane['mmap_threshold'])}
        result=subprocess.run(argv,cwd=REPO,env=os.environ|env,capture_output=True,text=True);(directory/'workload.log').write_text(result.stdout+result.stderr)
        after=custody.sources();assert before==after and result.returncode==0,result.stderr
        oracle=subprocess.run([sys.executable,'-B',str(ROOT/'verify-report.py'),str(report)],capture_output=True,text=True);(directory/'oracle.log').write_text(oracle.stdout+oracle.stderr);assert oracle.returncode==0,oracle.stderr
        row={'lane':i,**lane,'argv':argv,'cwd':str(REPO),'environment':env,'source_before':before,'source_after':after,'binary':binary,'exit_code':result.returncode,'oracle_exit_code':oracle.returncode,'artifacts':[{'path':str(p.relative_to(ROOT)),'sha256':sha(p),'bytes':p.stat().st_size} for p in sorted(directory.iterdir())]}
        rows.append(row);print(json.dumps({'lane':i,'status':'pass'}),flush=True)
    (ROOT/'allocator-regime-proof.json').write_text(json.dumps({'status':'pass','protocol_sha256':sha(ROOT/'allocator-regime-protocol.json'),'driver_sha256':sha(Path(__file__)),'glibc':subprocess.check_output(['getconf','GNU_LIBC_VERSION'],text=True).strip(),'rows':rows},indent=2)+'\n')
if __name__=='__main__':main()
