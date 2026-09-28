"""Root-only exact scanner disassembly before cleanup or workload comparison."""
import subprocess
import sys
import time
import custody as c

leg=sys.argv[1]
assert leg in ('before','after')
out=c.P/f'codegen-{leg}'
assert not out.exists()
build=c.read(c.P/f'build-{leg}/build.json')
source=c.read(c.P/f'build-{leg}/source.json')
assert c.source()==source
out.mkdir()
rows=[]
for variant in ('native','profile'):
    binary=build['binaries'][variant]
    assert c.artifact(binary['path'])==binary
    command=['nm','-S','--defined-only',binary['path']]
    started=time.time();result=subprocess.run(command,capture_output=True,text=True,check=True)
    assert not result.stderr
    matches=[line for line in result.stdout.splitlines() if '18scan_processed_xml17h' in line]
    assert len(matches)==1,matches
    addr,size,kind,symbol=matches[0].split();address,size=int(addr,16),int(size,16)
    assert kind.lower()=='t' and size>0
    symbols=out/f'{variant}.symbol.txt';symbols.write_text(matches[0]+'\n')
    nm={'command':command,'started':started,'ended':time.time(),'exit_code':result.returncode,
        'matched_count':len(matches),'symbols':c.artifact(symbols)}
    command=['objdump','-d','--demangle','--no-show-raw-insn',f'--start-address=0x{address:x}',
             f'--stop-address=0x{address+size:x}',binary['path']]
    started=time.time();result=subprocess.run(command,capture_output=True,text=True,check=True)
    assert not result.stderr
    assembly=out/f'{variant}.assembly.txt';assembly.write_text(result.stdout)
    rows.append({'variant':variant,'binary':binary,'source':build['source'],'nm':nm,
        'objdump':{'command':command,'started':started,'ended':time.time(),'exit_code':result.returncode,
                   'assembly':c.artifact(assembly)},'address':address,'size':size,'symbol':symbol})
    assert c.artifact(binary['path'])==binary and c.source()==source
versions={}
for tool in ('nm','objdump'):
    result=subprocess.run([tool,'--version'],capture_output=True,text=True,check=True)
    versions[tool]={'command':[tool,'--version'],'stdout':result.stdout,'stderr':result.stderr,'exit_code':result.returncode}
c.write(out/'receipt.json',{'schema':'litchi.performance.0813.codegen.v1','leg':leg,'rows':rows,'tools':versions,
    'scope':'Exact per-leg ordinary/profile scanner assembly; no causal cycle share or adoption claim.'})
print('0813',leg,'ordinary/profile scanner assembly retained')
