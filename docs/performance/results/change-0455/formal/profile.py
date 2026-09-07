#!/usr/bin/env python3
"""Separate whole-process hardware counters; never mixed with API timing lanes."""
import hashlib,json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[4]
EVENTS='cycles,instructions,branches,branch-misses,cache-misses,page-faults'
records=[]
for i,kind in enumerate(['baseline','candidate','candidate','baseline']):
    build=json.loads((ROOT/(kind+'-build.json')).read_text());binary=build['binaries']['normal'];assert hashlib.sha256(Path(binary['path']).read_bytes()).hexdigest()==binary['sha256']
    directory=ROOT/'profiles'/str(i);directory.mkdir(parents=True,exist_ok=False);report=directory/'report.json'
    argv=['taskset','-c','2','perf','stat','-x,','-e',EVENTS,'-o',str(directory/'perf.csv'),binary['path'],'provider-lifecycle','--corpus','media-rich','--provider','bytes','--samples','30','--warmup','3','--source-revision',build['revision'],'--output',str(report)]
    r=subprocess.run(argv,cwd=REPO,capture_output=True,text=True);(directory/'workload.log').write_text(r.stdout+r.stderr);assert r.returncode==0,r.stderr
    result=subprocess.run([sys.executable,'-B',str(ROOT/'verify-report.py'),str(report)],capture_output=True,text=True);(directory/'oracle.log').write_text(result.stdout+result.stderr);assert result.returncode==0,result.stderr
    records.append({'build':kind,'argv':argv,'binary':binary,'exit_code':r.returncode,'oracle_exit_code':result.returncode,'artifacts':[{'path':str(p.relative_to(ROOT)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size} for p in sorted(directory.iterdir())]})
(ROOT/'profile-proof.json').write_text(json.dumps({'status':'pass','scope':'separate whole-process counters including corpus generation, untimed output oracles and serialization; not API-attributed counters or ordinary timing samples','events':EVENTS,'order':['baseline','candidate','candidate','baseline'],'rows':records},indent=2)+'\n');print(json.dumps({'status':'pass','profiles':4}))
