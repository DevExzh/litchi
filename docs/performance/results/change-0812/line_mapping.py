"""Record DWARF inline attribution for the freshly sampled copy chain."""
import json
import subprocess
import time
import inputs as c

p=c.P/'fresh';out=p/'line-mapping.json';assert not out.exists()
complete=c.read(p/'complete.json');binary=complete['binary']
assert c.artifact(binary['path'])==binary
command=['addr2line','-afiC','-e',binary['path'],hex(complete['address']+0x24d),hex(complete['address']+0x27d)]
started=time.time();result=subprocess.run(command,capture_output=True,text=True,check=True)
version=subprocess.run(['addr2line','--version'],capture_output=True,text=True,check=True)
assert not result.stderr
out.write_text(json.dumps({'schema':'litchi.performance.0812.line-mapping.v1','binary':binary,
    'command':command,'started':started,'ended':time.time(),'exit_code':result.returncode,
    'stdout':result.stdout,'stderr':result.stderr,'tool_version':version.stdout,
    'scope':'DWARF source attribution for the fresh exact binary; no historical mapping or causal cost.'},indent=2,sort_keys=True)+'\n')
print(result.stdout)
