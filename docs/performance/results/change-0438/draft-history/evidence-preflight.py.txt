#!/usr/bin/env python3
"""Check frozen protocol and outer adapters against all preparatory reports."""
import importlib.util
from pathlib import Path
import json
ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('verify438',ROOT/'verify.py')
v=importlib.util.module_from_spec(spec);spec.loader.exec_module(v)
p=json.loads((ROOT/'protocol.json').read_text());v.validate_protocol(p)
o=v.load_oracle(p)
for mode in ['normal','allocator']:
 for shape in ['tiny','medium','large']:
  ids=[]
  for role in ['before-streaming','after-streaming']:
   path=ROOT/f'pilots/{role}/initial/{mode}-{shape}.json'
   value=o.validate_report(path,mode,shape,'after-streaming',samples=3,warmups=1)
   ids.append(v.normalized_identity(value,path))
  assert ids[0]==ids[1],(mode,shape)
print('PASS: 12 pilot oracle/outer adapters and same-API exact identities')
