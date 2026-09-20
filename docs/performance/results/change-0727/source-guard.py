#!/usr/bin/env python3
"""Verify exact restored baseline and the unchanged rejected candidate ancestry."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
a=read(P/'source-baseline.json');b=read(P/'source-candidate.json');rel='crates/litchi-xls/src/workbook/source.rs'
assert set(a)==set(b) and [k for k in a if a[k]!=b[k]]==[rel]
actual={str(p.relative_to(ROOT)):sha(p) for c in ['litchi-cfb','litchi-xls'] for p in (ROOT/'crates'/c).rglob('*') if p.is_file() and p.suffix in ('.rs','.toml')};assert actual==a
assert (P/'sources/baseline'/rel).read_bytes()==subprocess.check_output(['git','show',read(P/'plan.json')['baseline_revision']+':'+rel],cwd=ROOT)
assert (P/'sources/candidate'/rel).read_bytes()==(P.parent/'change-0726/candidate-source'/rel).read_bytes()
assert read(P/'build-restoration.json')==dict(exact=True,source_sha256=a)
for rel,h in read(P/'constraints.json').items():assert sha(ROOT/rel)==h
print('PASS exact baseline, candidate ancestry and all unchanged constraints')
