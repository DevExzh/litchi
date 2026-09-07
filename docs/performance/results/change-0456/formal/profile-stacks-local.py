#!/usr/bin/env python3
"""Paired diagnostic instruction profiles for the direct-payload memory experiment."""
import hashlib,json,os,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[4];TASK=Path('/tmp/litchi-goal-0456/stacks-local');TASK.mkdir()
rows=[]
for kind in ['baseline','candidate']:
    build=json.loads((ROOT/(kind+'-build.json')).read_text());binary=build['binaries']['normal'];assert hashlib.sha256(Path(binary['path']).read_bytes()).hexdigest()==binary['sha256']
    directory=ROOT/'stack-profiles-local'/kind;directory.mkdir(parents=True);data=TASK/(kind+'.data');report=directory/'report.json'
    argv=['taskset','-c','2','perf','record','--no-buildid-cache','-e','instructions','-c','50000000','--call-graph','dwarf,8192','-o',str(data),'--',binary['path'],'provider-lifecycle','--corpus','media-rich','--provider','bytes','--samples','30','--warmup','3','--source-revision',build['revision'],'--output',str(report)]
    result=subprocess.run(argv,cwd=REPO,env=os.environ | {'DEBUGINFOD_URLS':''},capture_output=True,text=True);(directory/'record.log').write_text(result.stdout+result.stderr);assert result.returncode==0,result.stderr
    for name,option in [('self','--no-children'),('inclusive','--children')]:
        r=subprocess.run(['perf','report','--stdio','--stdio-color','never',option,'--percent-limit','0.5','-i',str(data)],env=os.environ | {'DEBUGINFOD_URLS':''},capture_output=True,text=True);(directory/(name+'.txt')).write_text(r.stdout+r.stderr);assert r.returncode==0,r.stderr
    r=subprocess.run([sys.executable,'-B',str(ROOT/'verify-report.py'),str(report)],capture_output=True,text=True);(directory/'oracle.log').write_text(r.stdout+r.stderr);assert r.returncode==0,r.stderr
    rows.append({'build':kind,'argv':argv,'binary':binary,'exit_code':result.returncode,'oracle_exit_code':r.returncode,'raw_profile':{'path':str(data),'sha256':hashlib.sha256(data.read_bytes()).hexdigest(),'bytes':data.stat().st_size},'artifacts':[{'path':str(p.relative_to(ROOT)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size} for p in sorted(directory.iterdir())]})
(ROOT/'stack-profile-local-proof.json').write_text(json.dumps({'status':'pass','scope':'one baseline and candidate whole-process instruction-sampling diagnostic including user and kernel events; local symbols only, DEBUGINFOD_URLS empty; includes corpus/oracles, not API-only samples or timing comparison','rows':rows},indent=2)+'\n');print(json.dumps({'status':'pass','profiles':2}))
