#!/usr/bin/env python3
"""Retain assembly of the measured attribute getter with executable custody."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
ROOT=Path(__file__).resolve().parent
variant=sys.argv[1];assert variant in ["baseline","candidate"]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
binding_path=ROOT/(variant+"-binding.json");binding=json.loads(binding_path.read_text());binary=Path(binding["binaries"]["normal"]["path"])
assert sha(binary)==binding["binaries"]["normal"]["sha256"]
commands=[];artifacts=[]
for method in ["get","lookup"]:
 symbol="<litchi_odp::codec::parser::codec::xml::validation::ElementAttrs>::"+method
 argv=["objdump","-d","--demangle","--no-show-raw-insn","--disassemble="+symbol,str(binary)]
 raw=subprocess.check_output(argv);path=ROOT/(variant+"-"+method+".asm.txt")
 if path.exists():assert path.read_bytes()==raw
 else:
  with path.open("xb") as stream:stream.write(raw)
 commands.append(argv);artifacts.append({"path":path.name,"bytes":len(raw),"sha256":sha(path)})
assert sha(binary)==binding["binaries"]["normal"]["sha256"]
with (ROOT/(variant+"-assembly.json")).open("x") as stream:
 json.dump({"change":461,"variant":variant,"binding_sha256":sha(binding_path),"binary_sha256":sha(binary),"commands":commands,"artifacts":artifacts},stream,indent=2);stream.write("\n")
print(variant,"assembly bound")
