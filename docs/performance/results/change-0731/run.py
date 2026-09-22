#!/usr/bin/env python3
"""Serial qualification, frozen collection and offline replay for the PPT owner."""
import hashlib,json,os,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];TARGET=ROOT.parent/'litchi-target-0731';BIN=ROOT.parent/'litchi-0731-bin'
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def source():return {str(p.relative_to(ROOT)):sha(p) for p in sorted(list((ROOT/'crates').rglob('*.rs'))+list((ROOT/'crates').rglob('Cargo.toml'))+[ROOT/'Cargo.toml',ROOT/'Cargo.lock'])}
def execute(cmd,out,env=None):
 start=time.monotonic()
 with out.open('wb') as f:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
 return dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,output=str(out.relative_to(P)),sha256=sha(out))
def guard():
 assert all(sha(ROOT/f)==h for f,h in read(P/'constraints.json').items())
 if (P/'build.json').exists():assert source()==read(P/'build.json')['source']
def bindings():
 paths=[P/n for n in ['run.py','plan.json','hypothesis.md','constraints.json','case.json','build.json','oracle.json','analyze.py','environment.json','ancestry.json']]
 paths+=list((P/'probe').rglob('*'))+list((P/'quality').rglob('*'))
 return {str(p):sha(p) for p in paths if p.is_file()}
mode=sys.argv[1];guard()
if mode=='build':
 out=P/'quality';out.mkdir();before=source();env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2',RUSTDOCFLAGS='-D warnings');m=str(P/'probe/Cargo.toml');common=['--manifest-path',m,'--release','--offline','--locked'];commands=[['cargo','fmt','--manifest-path',m,'--','--check'],['cargo','test',*common,'--lib'],['cargo','clippy',*common,'--all-targets','--','-D','warnings'],['cargo','doc',*common,'--no-deps'],['cargo','build',*common,'--bins']];runs=[]
 for i,cmd in enumerate(commands):
  row=execute(cmd,out/f'{i}.log',env);runs.append(row);write(out/'manifest.json',runs);print(i,row['exit_code'],flush=True);assert row['exit_code']==0
 assert source()==before;BIN.mkdir();binaries=[]
 for name in ['ole_format_save_probe','ole_format_save_probe_alloc']:
  dest=BIN/name;shutil.copy2(TARGET/'release'/name,dest);binaries.append(dict(path=str(dest),sha256=sha(dest),bytes=dest.stat().st_size))
 write(P/'build.json',dict(source=before,binaries=binaries));print('PASS build and five probe gates')
elif mode=='freeze':
 assert not (P/'freeze.json').exists();write(P/'freeze.json',bindings());print('frozen')
elif mode=='capture':
 assert read(P/'freeze.json')==bindings();out=P/'captures';out.mkdir();plan=read(P/'plan.json');case=read(P/'case.json');rows=[]
 for repeat in range(3):
  for lane in ['native','allocation','callgrind']:
   name=f'{lane}-{repeat}';binary=BIN/('ole_format_save_probe_alloc' if lane=='allocation' else 'ole_format_save_probe');identity=next(x for x in read(P/'build.json')['binaries'] if x['path']==str(binary));assert sha(binary)==identity['sha256'];cmd=['taskset','-c','12']
   if lane=='callgrind':cmd+=['valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect='+plan['callgrind']['symbol'],'--callgrind-out-file='+str(out/(name+'.callgrind')),'--log-file='+str(out/(name+'.valgrind'))]
   cmd+=[str(binary),'--case','ppt45543','--input',case['path'],'--operation','format','--samples',str(50 if lane=='native' else 1),'--warmups',str(3 if lane=='native' else 0)]
   row=execute(cmd,out/(name+'.json'));row.update(lane=lane,repeat=repeat);rows.append(row);write(out/'manifest.json',dict(status='running',freeze_sha256=sha(P/'freeze.json'),runs=rows));print(name,row['exit_code'],flush=True);assert row['exit_code']==0
   if lane=='callgrind':
    for suffix in ['callgrind','valgrind']:row[suffix+'_sha256']=sha(out/(name+'.'+suffix))
    cmd=['callgrind_annotate','--inclusive=yes','--tree=both','--threshold=99',str(out/(name+'.callgrind'))];annotation=execute(cmd,out/(name+'.annotation'));assert annotation['exit_code']==0;row['annotation']=annotation
 assert read(P/'freeze.json')==bindings();guard();write(out/'manifest.json',dict(status='complete',freeze_sha256=sha(P/'freeze.json'),runs=rows))
else:raise SystemExit('build|freeze|capture')
