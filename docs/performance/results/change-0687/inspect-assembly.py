#!/usr/bin/env python3
"""Retain disassembly of the measured private helper from both actual binaries."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;rows=[]
for phase,suffix in [('baseline','before'),('candidate','after')]:
 binary=Path('/home/zhuhe/code/litchi-target-0687-'+suffix)/'release/xls0684-repeat';symbols=subprocess.check_output(['nm','-S','--defined-only',str(binary)],text=True)
 selected=[line.split() for line in symbols.splitlines() if 'directory_name19directory_name_data' in line];assert len(selected)==1,selected
 address,size,kind,symbol=selected[0];cmd=['objdump','-d','-C','--start-address=0x'+address,'--stop-address='+hex(int(address,16)+int(size,16)),str(binary)]
 output=subprocess.check_output(cmd,text=True);f=P/(phase+'-directory-name-assembly.txt');f.write_text(output)
 rows.append(dict(phase=phase,command=cmd,symbol=symbol,code_bytes=int(size,16),binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest(),assembly_sha256=hashlib.sha256(f.read_bytes()).hexdigest()))
(P/'assembly-manifest.json').write_text(json.dumps(rows,indent=2)+'\n')
