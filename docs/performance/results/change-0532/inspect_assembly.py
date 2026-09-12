"""Record symbol-bounded generated code for the three residual CFB owners."""
import json
from run import HERE,FOLDER,SCRATCH,run,sha,write
binary=SCRATCH/'normal'
owners=('claim_sector','validate_stream_allocations','validate_physical_sector_layout')
write(HERE/'assembly-plan.json',dict(schema='litchi-0532-assembly-plan-v1',status='frozen-before-inspection',owners=list(owners),binary_sha256=sha(binary),source_manifest_sha256=sha(FOLDER/'source-manifest.json'),script_sha256=sha(HERE/'inspect_assembly.py'),scope='Disassembly explains measured code shape; static instructions are not a latency or dynamic-count claim.'))
run('symbols',['nm','-S','--defined-only',str(binary)],binary)
rows=[]
for line in (FOLDER/'symbols.stdout').read_text().splitlines():
    fields=line.split()
    if len(fields)!=4:continue
    address,size,kind,symbol=fields
    if kind not in ('t','T'):continue
    owner=next((o for o in owners if o in symbol),None)
    if owner is None:continue
    name='assembly-'+owner+'-'+str(sum(r['owner']==owner for r in rows))
    command=['objdump','-d','--disassemble='+symbol,str(binary)]
    run(name,command,binary)
    rows.append(dict(name=name,owner=owner,symbol=symbol,address_hex=address,size_bytes=int(size,16),receipt_sha256=sha(FOLDER/(name+'.receipt.json'))))
assert set(r['owner'] for r in rows)==set(owners)
write(HERE/'assembly-index.json',dict(schema='litchi-0532-assembly-index-v1',plan_sha256=sha(HERE/'assembly-plan.json'),symbols_receipt_sha256=sha(FOLDER/'symbols.receipt.json'),rows=rows))
