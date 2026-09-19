#!/usr/bin/env python3
"""Recheck documentation classifications after final report edits."""
import json,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];out=P/'evidence';out.mkdir(exist_ok=True);results=[]
for previous in json.loads((P.parent/'change-0687/evidence/final-doc-results.json').read_text()):
 name=previous['name'];cmd=previous['command'];start=time.monotonic()
 with (out/('final-doc-'+name+'.log')).open('w') as log:r=subprocess.run(cmd,cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
 results.append(dict(name=name,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start));(out/'final-doc-results.json').write_text(json.dumps(results,indent=2)+'\n');print(name,r.returncode,flush=True);assert r.returncode==0
