"""Finish only post-capture assembly after the retained fresh-driver NameError."""
import json
import subprocess
import time
from pathlib import Path
import inputs as c

O=c.P/'fresh'
assert not (O/'complete.json').exists()
frozen=c.read(O/'frozen.json');binary=frozen['binary']
assert c.artifact(binary['path'])==binary
assert c.verify()==frozen['inputs']
for row in frozen['drivers'].values():assert c.artifact(row['path'])==row
captures=c.read(O/'receipts.json');decodes=c.read(O/'decode.json')
assert len(captures)==len(decodes)==2
for row in captures:
    assert row['exit_code']==0 and c.artifact(row['report']['path'])==row['report']
for row in decodes:assert row['exit_code']==0
for row in c.read(O/'compression.json'):
    assert c.artifact(row['compressed']['path'])==row['compressed']
assert (O/'symbols.txt').exists() and not (O/'nm.json').exists()
def write(name,value):
    path=O/name;assert not path.exists()
    path.write_text(json.dumps(value,indent=2,sort_keys=True)+'\n')
def command(args,path):
    started=time.time()
    result=subprocess.run(args,cwd=c.ROOT,capture_output=True,text=True)
    assert not path.exists();path.write_text(result.stdout)
    row={'command':args,'started':started,'ended':time.time(),'exit_code':result.returncode,
         'output':c.artifact(path),'stderr':result.stderr,'binary':binary}
    assert result.returncode==0 and not result.stderr
    return row
write('assembly-resume.json',{'schema':'litchi.performance.0812.assembly-resume.v1',
    'reason':'fresh.py completed two captures, two decodes and compression, then NameError: nm is not defined (local variable named m).',
    'workload_retries':0,'prior_nm_stdout':c.artifact(O/'symbols.txt'),
    'driver':c.artifact(c.P/'finish_assembly.py'),'started':time.time()})
row=command(['nm','-S','--defined-only',binary['path']],O/'symbols-resumed.txt');write('nm.json',row)
assert (O/'symbols-resumed.txt').read_bytes()==(O/'symbols.txt').read_bytes()
lines=[line for line in (O/'symbols-resumed.txt').read_text().splitlines() if '18scan_processed_xml17h' in line]
assert len(lines)==1
address,size,kind,symbol=lines[0].split();base,size=int(address,16),int(size,16)
assert kind.lower()=='t'
args=['objdump','-d','--demangle','--no-show-raw-insn',f'--start-address=0x{base:x}',f'--stop-address=0x{base+size:x}',binary['path']]
row=command(args,O/'scanner-assembly.txt');write('objdump.json',row)
assert c.artifact(binary['path'])==binary and c.verify()==frozen['inputs']
write('complete.json',{'schema':'litchi.performance.0812.fresh-complete.v1','binary':binary,
    'reports':2,'samples':200,'inputs':frozen['inputs'],'address':base,'size':size,'symbol':symbol,
    'assembly':c.artifact(O/'scanner-assembly.txt'),'ended':time.time()})
print('0812 post-capture assembly completion PASS; zero workload retries')
