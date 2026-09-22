#!/usr/bin/env python3
"""Verify all source bytes used by the baseline build remain unchanged."""
import hashlib,json
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
x=json.loads((P/'builds.json').read_text());expected=x['source_sha256']
paths=[p for p in (ROOT/'crates').rglob('*') if p.is_file() and p.suffix in ('.rs','.toml')]+[ROOT/'Cargo.toml',ROOT/'Cargo.lock']
assert {str(p.relative_to(ROOT)):sha(p) for p in paths}==expected
for rel,h in json.loads((P/'constraints.json').read_text()).items():assert sha(ROOT/rel)==h
assert {str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()}==x['probe_sha256']
print('PASS complete workspace source, probe and constraint identities')
