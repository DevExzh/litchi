#!/usr/bin/env python3
"""Retain static symbol and section tables for the two diagnostic binaries."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
out=P/'symbols';out.mkdir(exist_ok=True)
records=[]
for phase in ['baseline','candidate']:
 build=json.loads((P/f'build-{phase}.json').read_text());binary=build['binary']
 assert hashlib.sha256(Path(binary).read_bytes()).hexdigest()==build['binary_sha256']
 for name,command in [('nm',['nm','-S','-C',binary]),('size',['size','-A',binary])]:
  result=subprocess.run(command,capture_output=True);assert result.returncode==0
  output=out/f'{phase}-{name}.txt';output.write_bytes(result.stdout)
  records.append(dict(phase=phase,name=name,command=command,exit_code=result.returncode,binary_sha256=build['binary_sha256'],stdout=str(output.relative_to(P)),stdout_sha256=hashlib.sha256(result.stdout).hexdigest(),stderr=result.stderr.decode()))
(P/'symbols.json').write_text(json.dumps(records,indent=2)+'\n')
