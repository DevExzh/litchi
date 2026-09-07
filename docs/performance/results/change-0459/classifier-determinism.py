#!/usr/bin/env python3
"""Require identical classification under independently seeded Python hashes."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
root=Path(__file__).resolve().parent
outputs=[]
for seed in ["1","2718"]:
 result=subprocess.run([sys.executable,"-B",str(root/"classify.py")],cwd=root,env=os.environ|{"PYTHONHASHSEED":seed,"PYTHONDONTWRITEBYTECODE":"1"},capture_output=True,check=True)
 outputs.append(result.stdout)
assert outputs[0]==outputs[1]==(root/"diagnostic-summary.json").read_bytes()
print(json.dumps({"status":"pass","hash_seeds":[1,2718],"summary_sha256":hashlib.sha256(outputs[0]).hexdigest()}))
