"""Root-only exact scanner/inspector disassembly before cleanup or workload comparison."""
import subprocess
import sys
import time
import custody as c

out=c.P/'assembly'
assert not out.exists()
build=c.read(c.P/'build/build.json')
source=c.read(c.P/'build/source.json')
assert c.source()==source
out.mkdir()
rows=[]
for variant in ('control','profile','fp'):
    binary=build['binaries'][variant]
    assert c.artifact(binary['path'])==binary
    nm_command=['nm','-S','--defined-only',binary['path']]
    nm_started=time.time();nm_result=subprocess.run(nm_command,capture_output=True,text=True,check=True)
    nm_ended=time.time()
    assert not nm_result.stderr
    for name,pattern in (('scanner','18scan_processed_xml17h'),('inspector','15inspect_element17h')):
        matches=[line for line in nm_result.stdout.splitlines() if pattern in line]
        assert len(matches)==1,matches
        addr,size,kind,symbol=matches[0].split();address,size=int(addr,16),int(size,16)
        assert kind.lower()=='t' and size>0
        symbols=out/f'{variant}.{name}.symbol.txt';symbols.write_text(matches[0]+'\n')
        nm={'command':nm_command,'started':nm_started,'ended':nm_ended,'exit_code':nm_result.returncode,
            'matched_count':len(matches),'symbols':c.artifact(symbols)}
        command=['objdump','-d','--demangle','--no-show-raw-insn',f'--start-address=0x{address:x}',
                 f'--stop-address=0x{address+size:x}',binary['path']]
        started=time.time();result=subprocess.run(command,capture_output=True,text=True,check=True)
        assert not result.stderr
        assembly=out/f'{variant}.{name}.assembly.txt';assembly.write_text(result.stdout)
        rows.append({'variant':variant,'name':name,'binary':binary,'source':build['source'],'nm':nm,
            'objdump':{'command':command,'started':started,'ended':time.time(),'exit_code':result.returncode,
                       'assembly':c.artifact(assembly)},'address':address,'size':size,'symbol':symbol})
    assert c.artifact(binary['path'])==binary and c.source()==source
versions={}
for tool in ('nm','objdump'):
    result=subprocess.run([tool,'--version'],capture_output=True,text=True,check=True)
    versions[tool]={'command':[tool,'--version'],'stdout':result.stdout,'stderr':result.stderr,'exit_code':result.returncode}
c.write(out/'receipt.json',{'schema':'litchi.performance.0814.codegen.v1','rows':rows,'tools':versions,
    'scope':'Exact current control/profile/fp scanner/inspector assembly; no causal cycle share or adoption claim.'})
print('0814 exact control/profile/fp scanner/inspector assembly retained')
