"""Capture current collector and bitset symbol-bounded disassembly."""
import run as R

def inspect(stage):
    R.configure(stage)
    binary = R.SCRATCH / 'normal'
    plan = __import__('json').loads((R.HERE / 'plan.json').read_text())
    R.run('symbols',['nm','-S','--defined-only',str(binary)],binary)
    rows=[]
    for line in (R.FOLDER/'symbols.stdout').read_text().splitlines():
        fields=line.split()
        if len(fields)!=4:continue
        address,size,kind,symbol=fields
        if kind not in ('t','T') or not any(owner in symbol for owner in plan['assembly']['owners']):continue
        name='assembly-'+str(len(rows))
        R.run(name,['objdump','-d','--disassemble='+symbol,str(binary)],binary)
        rows.append(dict(name=name,symbol=symbol,address_hex=address,size_bytes=int(size,16),receipt_sha256=R.sha(R.FOLDER/(name+'.receipt.json'))))
    assert all(any(name in row['symbol'] for row in rows) for name in plan['assembly']['required'])
    R.write(R.FOLDER/'assembly-index.json',dict(plan_sha256=R.sha(R.HERE/'plan.json'),binary_sha256=R.sha(binary),source_manifest_sha256=R.sha(R.FOLDER/'source-manifest.json'),script_sha256=R.sha(R.HERE/'inspect_assembly.py'),rows=rows))
