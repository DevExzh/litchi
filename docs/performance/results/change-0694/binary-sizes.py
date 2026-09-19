#!/usr/bin/env python3
"""Price native code size separately from debug-bearing binary file length."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
records={}
for phase in ['baseline','candidate']:
 records[phase]={}
 for label in ['native','allocations','refusal']:
  binary=ROOT.parent/'litchi-0694-bin'/f'{phase}-{label}'
  command=['size',str(binary)]
  result=subprocess.run(command,capture_output=True,text=True,check=True)
  records[phase][label]=dict(bytes=binary.stat().st_size,sha256=hashlib.sha256(binary.read_bytes()).hexdigest(),size_command=command,exit_code=result.returncode,size_output=result.stdout)
(P/'binary-sizes.json').write_text(json.dumps(records,indent=2)+'\n')
