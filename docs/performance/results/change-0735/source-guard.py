#!/usr/bin/env python3
"""Recheck byte-identical typed-validation prefixes and their retained proof."""
import hashlib,json
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
f=json.loads((P/'base.json').read_text())['owned_file'];before=(P/'source-archive/before'/f).read_text();after=(ROOT/f).read_text();rows=[]
for name in ['replace_persisted_record','insert_persisted_record']:
 a=before[before.index('pub(crate) fn '+name):];b=after[after.index('pub(crate) fn '+name):];ap=a[:a.index('    let mut candidate')].rstrip();bp=b[:b.index('    //')].rstrip();assert ap==bp,name
 rows.append(dict(function=name,validation_prefix_bytes=len(ap.encode()),validation_prefix_sha256=hashlib.sha256(ap.encode()).hexdigest(),byte_identical=True))
x=json.loads((P/'source-equivalence.json').read_text());assert x['status']=='passed' and x['functions']==rows
assert x['before_sha256']==hashlib.sha256(before.encode()).hexdigest() and x['after_sha256']==hashlib.sha256(after.encode()).hexdigest()
assert all(hashlib.sha256((ROOT/f).read_bytes()).hexdigest()==h for f,h in json.loads((P/'constraints.json').read_text()).items())
print('PASS unchanged typed-validation prefixes, source identities and accepted constraints')
