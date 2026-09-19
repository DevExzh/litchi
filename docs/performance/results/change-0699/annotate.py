#!/usr/bin/env python3
"""Bind instruction samples and bounded disassembly to the marker scan symbol."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
out=P/'annotations';out.mkdir(exist_ok=True)
symbol='litchi_ooxml_common::mce::codec::process_markup_compatibility'
records=[]
for phase in ['baseline','candidate']:
 build=json.loads((P/f'build-{phase}.json').read_text());binary=Path(build['binary'])
 data=ROOT.parent/'litchi-0699-profile'/f'{phase}-early-name-error.data'
 sha=lambda f:hashlib.sha256(f.read_bytes()).hexdigest()
 assert sha(binary)==build['binary_sha256']
 rows=[line.split(maxsplit=3) for line in (P/f'symbols/{phase}-nm.txt').read_text().splitlines()]
 row=next(r for r in rows if len(r)==4 and r[3]==symbol)
 start=int(row[0],16);size=int(row[1],16)
 commands=[('assembly',['objdump','-d','-C','--no-show-raw-insn',f'--start-address={start}',f'--stop-address={start+size}',str(binary)]),('samples',['perf','annotate','--stdio','--show-nr-samples','--symbol',symbol,'-i',str(data)])]
 for name,command in commands:
  result=subprocess.run(command,capture_output=True)
  stdout=out/f'{phase}-{name}.stdout';stderr=out/f'{phase}-{name}.stderr'
  stdout.write_bytes(result.stdout);stderr.write_bytes(result.stderr)
  records.append(dict(phase=phase,name=name,command=command,exit_code=result.returncode,symbol=symbol,address=start,size=size,binary_sha256=sha(binary),raw_data_sha256=sha(data),stdout_sha256=sha(stdout),stderr_sha256=sha(stderr)))
  (P/'annotations.json').write_text(json.dumps(records,indent=2)+'\n')
  assert result.returncode==0,result.stderr.decode()
print('Four marker-scan instruction artifacts retained.')
