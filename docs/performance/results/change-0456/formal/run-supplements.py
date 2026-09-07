#!/usr/bin/env python3
"""Run declared native and hardware-counter supplements serially."""
import hashlib,json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
protocol=json.loads((ROOT/'supplement-protocol.json').read_text())
for name,digest in protocol['bound_files'].items():assert hashlib.sha256((ROOT/name).read_bytes()).hexdigest()==digest
for name,argv in protocol['commands'].items():
    subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',name,'--',*argv],check=True)
