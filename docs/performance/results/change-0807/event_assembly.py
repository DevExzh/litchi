"""Post-capture exploratory disassembly of the leading sampled event leaf."""
import subprocess,time
import custody as c
receipt=c.P/'event-assembly.json';assert not receipt.exists()
binary=c.read(c.P/'build-fp/receipt.json')['binary']
assert c.artifact(binary['path'])==binary
symbol_file=c.P/'event-symbols.txt';assembly_file=c.P/'event-assembly.txt'
cmd=['nm','-S','--defined-only',binary['path']]
start=time.time();r=subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
assert r.returncode==0 and not r.stderr
symbol_file.write_text(r.stdout)
rows=[line.split() for line in r.stdout.splitlines() if 'NsReader' in line and '13process_event' in line]
assert len(rows)==1 and len(rows[0])==4
address,size,kind,symbol=rows[0];assert kind.lower()=='t'
address=int(address,16);size=int(size,16);assert size>0
command=['objdump','-d','--demangle','--no-show-raw-insn',f'--start-address={address}',f'--stop-address={address+size}',binary['path']]
d=subprocess.run(command,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True)
assert d.returncode==0 and not d.stderr
assembly_file.write_text(d.stdout)
c.write(receipt,{'scope':'Post-capture exploratory assembly only; no instruction frequency, causal latency fraction or optimization claim.','binary':binary,'symbol':symbol,'address':address,'size':size,'started':start,'ended':time.time(),'nm_command':cmd,'nm_exit_code':r.returncode,'symbols':c.artifact(symbol_file),'objdump_command':command,'objdump_exit_code':d.returncode,'assembly':c.artifact(assembly_file)})
print('Exploratory exact-symbol assembly retained',flush=True)
