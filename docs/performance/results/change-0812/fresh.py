"""Root-only fresh native samples after exact reconstruction was refused."""
import gzip
import io
import json
from pathlib import Path
import subprocess
import time
import inputs as c

O = c.P / 'fresh'
assert not O.exists()
r = c.read(c.P / 'rebuild/receipt.json')
assert r['status'] == 'binary-mismatch' and not r['historical_mapping_authorized']
binary = r['binary']
assert c.artifact(binary['path']) == binary
custody = c.verify()
O.mkdir()
def write(name,value):
    p=O/name
    assert not p.exists()
    p.write_text(json.dumps(value,indent=2,sort_keys=True)+'\n')
def stable():
    assert c.verify() == custody
    assert c.artifact(binary['path']) == binary

def run(command, output, errors=None):
    started=time.time()
    with output.open('w') as out:
        if errors:
            with errors.open('w') as err:
                result=subprocess.run(command,cwd=c.ROOT,stdout=out,stderr=err)
        else:
            result=subprocess.run(command,cwd=c.ROOT,stdout=out,stderr=subprocess.STDOUT)
    row={'command':command,'started':started,'ended':time.time(),'exit_code':result.returncode,
         'output':c.artifact(output),'errors':c.artifact(errors) if errors else None,'binary':binary}
    return row

versions={}
for tool in ('perf','nm','objdump'):
    q=subprocess.run([tool,'--version'],capture_output=True,text=True,check=True)
    versions[tool]={'command':[tool,'--version'],'stdout':q.stdout,'stderr':q.stderr,'exit_code':q.returncode}
write('frozen.json',{'schema':'litchi.performance.0812.fresh-frozen.v1','inputs':custody,
    'binary':binary,'build':c.artifact(c.P/'rebuild/receipt.json'),'tool_versions':versions,
    'drivers':{n:c.artifact(c.P/n) for n in ('fresh.py','fresh-plan.md','inputs.py')}})
receipts=[]
for repeat in range(2):
    stable()
    raw,report,log=(O/f'{repeat}.{suffix}' for suffix in ('data','json','log'))
    cmd=['taskset','-c','12','perf','record','--no-buildid-cache','-e','cycles:u','-F','499',
         '--call-graph','fp','-o',str(raw),'--',binary['path'],'--mode','capture','--shape','large',
         '--samples','100','--warmup','0','--output',str(report)]
    row=run(cmd,log);row.update(repeat=repeat,raw=c.artifact(raw),report=c.artifact(report))
    write(f'{repeat}.receipt.json',row)
    assert row['exit_code']==0
    receipts.append(row);stable()
    print('fresh perf',repeat,'PASS',flush=True)
write('receipts.json',receipts)
decodes=[]
for row in receipts:
    stable(); repeat=row['repeat']
    raw=Path(row['raw']['path']);assert c.artifact(raw)==row['raw']
    frames=O/f'{repeat}.frames';errors=O/f'{repeat}.decode.log'
    decoded=run(['perf','script','--no-inline','--ns','-i',str(raw)],frames,errors)
    decoded.update(repeat=repeat,raw=row['raw'])
    write(f'{repeat}.decode.json',decoded)
    assert decoded['exit_code']==0 and frames.stat().st_size>0
    decodes.append(decoded);stable()
write('decode.json',decodes)
compressed=[]
for repeat in range(2):
    for suffix in ('data','frames'):
        path=O/f'{repeat}.{suffix}';before=c.artifact(path);raw=path.read_bytes()
        output=io.BytesIO()
        with gzip.GzipFile(filename='',mode='wb',fileobj=output,mtime=0,compresslevel=9) as gz:gz.write(raw)
        dest=path.with_suffix(path.suffix+'.gz');dest.write_bytes(output.getvalue())
        assert gzip.decompress(dest.read_bytes())==raw
        compressed.append({'repeat':repeat,'kind':suffix,'original':before,'compressed':c.artifact(dest)})
        path.unlink()
write('compression.json',compressed)
# This assembly belongs only to these fresh samples, never the old offsets.
m=run(['nm','-S','--defined-only',binary['path']],O/'symbols.txt')
write('nm.json',nm);assert nm['exit_code']==0
lines=[line for line in (O/'symbols.txt').read_text().splitlines() if '18scan_processed_xml17h' in line]
assert len(lines)==1
address,size,kind,symbol=lines[0].split();base,size=int(address,16),int(size,16)
assert kind.lower()=='t'
command=['objdump','-d','--demangle','--no-show-raw-insn',f'--start-address=0x{base:x}',f'--stop-address=0x{base+size:x}',binary['path']]
assembly=run(command,O/'scanner-assembly.txt');write('objdump.json',assembly)
assert assembly['exit_code']==0
stable()
write('complete.json',{'schema':'litchi.performance.0812.fresh-complete.v1','binary':binary,'reports':2,
    'samples':200,'inputs':custody,'address':base,'size':size,'symbol':symbol,
    'assembly':c.artifact(O/'scanner-assembly.txt'),'ended':time.time()})
print('fresh captures, exact-binary decode, compression and assembly PASS',flush=True)
