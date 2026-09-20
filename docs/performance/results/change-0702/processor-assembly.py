#!/usr/bin/env python3
"""Bind the marker-search processor's static assembly to each native binary."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
phase=sys.argv[1];assert phase in ['baseline','candidate']
build=next(r for r in json.loads((P/f'builds-{phase}.json').read_text()) if r['label']=='native')
binary=Path(build['binary']);sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
assert sha(binary)==build['binary_sha256']
out=P/'processor-assembly';out.mkdir(exist_ok=True)
commands=[]
def run(name,command):
 r=subprocess.run(command,capture_output=True);stdout=out/f'{phase}-{name}.stdout';stderr=out/f'{phase}-{name}.stderr'
 stdout.write_bytes(r.stdout);stderr.write_bytes(r.stderr)
 commands.append(dict(name=name,command=command,exit_code=r.returncode,stdout_sha256=sha(stdout),stderr_sha256=sha(stderr)))
 assert r.returncode==0,r.stderr.decode()
 return r.stdout.decode()
table=run('symbols',['nm','-S','-C',str(binary)])
symbol='litchi_ooxml_common::mce::codec::process_markup_compatibility'
rows=[line.split(maxsplit=3) for line in table.splitlines()]
row=next(r for r in rows if len(r)==4 and r[3]==symbol)
start=int(row[0],16);size=int(row[1],16)
run('processor',['objdump','-d','-C','--no-show-raw-insn',f'--start-address={start}',f'--stop-address={start+size}',str(binary)])
extra={}
helper_symbol='litchi_ooxml_common::mce::codec::contains_mce_namespace'
helper=next(r for r in rows if len(r)==4 and r[3]==helper_symbol)
helper_start=int(helper[0],16);helper_size=int(helper[1],16)
run('marker-search',['objdump','-d','-C','--no-show-raw-insn',f'--start-address={helper_start}',f'--stop-address={helper_start+helper_size}',str(binary)])
extra=dict(helper_symbol=helper_symbol,helper_address=helper_start,helper_size=helper_size)
(out/f'{phase}.json').write_text(json.dumps(dict(phase=phase,binary_sha256=sha(binary),symbol=symbol,address=start,size=size,commands=commands,**extra),indent=2)+'\n')
print(phase,'processor assembly retained')
