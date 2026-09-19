#!/usr/bin/env python3
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
binary=Path('/home/zhuhe/code/litchi-target-0686-after/release/xls0686-generate')
template='test-data/poi/test-data/spreadsheet/Simple.xls'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
commands=[];repeat=Path('/home/zhuhe/code/litchi-0686-profile/generator-repeat')
for directory in [P/'fixtures',repeat]:
 command=[str(binary),template,str(directory)];r=subprocess.run(command,cwd=ROOT,capture_output=True,text=True);assert r.returncode==0,r.stderr
 commands.append(dict(command=command,exit_code=r.returncode,stdout=r.stdout))
files={}
for n in [70000,100000]:
 f=P/'fixtures'/f'numeric-{n}.xls';assert sha(f)==sha(repeat/f.name)
 files[str(f.relative_to(ROOT))]=dict(sha256=sha(f),bytes=f.stat().st_size,first_sheet_occurrences=n+1)
manifest=dict(commands=commands,binary_sha256=sha(binary),template={template:sha(ROOT/template)},generator_sha256={str(f.relative_to(ROOT)):sha(f) for f in sorted((P/'generator').rglob('*')) if f.is_file()},files=files,identical_two_runs=True)
(P/'generator-manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
cases=json.loads((P/'cases.json').read_text());cases=[c for c in cases if not c['case'].startswith('synthetic-')]
for n in [70000,100000]:cases.append(dict(case=f'synthetic-{n}-default',path=str((P/'fixtures'/f'numeric-{n}.xls').relative_to(ROOT)),sheet=0,row=0,column=0,budget=2097152))
(P/'cases.json').write_text(json.dumps(cases,indent=2)+'\n')
print('Generated identical fixtures twice:',files)
