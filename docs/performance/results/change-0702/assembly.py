#!/usr/bin/env python3
"""Retain native start-function assembly for mechanism review."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
phase=sys.argv[1];assert phase in ['baseline','candidate']
binary=ROOT.parent/'litchi-0702-bin'/(phase+'-native')
resolved=subprocess.check_output(['nm','--defined-only','-C',str(binary)],text=True)
rows=[line.split() for line in resolved.splitlines() if line.endswith(' litchi_ooxml_common::mce::codec::start')]
assert len(rows)==1,rows
address=rows[0][0]
raw=subprocess.check_output(['nm','--defined-only',str(binary)],text=True)
symbol=next(line.split()[2] for line in raw.splitlines() if line.split()[0]==address and 'start' in line)
command=['objdump','-d','-l','--disassemble='+symbol,str(binary)]
r=subprocess.run(command,capture_output=True,text=True,check=True)
folder=P/'assembly';folder.mkdir(exist_ok=True)
output=folder/(phase+'-start.txt');output.write_text(r.stdout)
(folder/(phase+'.json')).write_text(json.dumps(dict(command=command,symbol=symbol,exit_code=r.returncode,binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest(),output_sha256=hashlib.sha256(output.read_bytes()).hexdigest(),stderr=r.stderr),indent=2)+'\n')
print(phase,'assembly retained')
