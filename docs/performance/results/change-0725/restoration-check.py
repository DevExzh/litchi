#!/usr/bin/env python3
"""Read-only final baseline, constraint, rejection and cleanup verification."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
for rel,digest in json.loads((P/'constraints.json').read_text()).items():assert sha(ROOT/rel)==digest,rel
r=json.loads((P/'restoration.json').read_text());assert r['exact'] and r['disposition']=='rejected'
actual={str(p.relative_to(ROOT)):sha(p) for name in ('litchi-cfb','litchi-xls') for p in (ROOT/'crates'/name).rglob('*.rs')}
assert actual==r['baseline_rust_source_sha256']
for row in r['paths']:
 assert sha(P/'candidate-source'/row['path'])==row['candidate_sha256']
 if row['removed']:assert not (ROOT/row['path']).exists()
 else:assert sha(ROOT/row['path'])==row['baseline_sha256']
subprocess.run(['git','diff','--exit-code',r['baseline_revision'],'--','crates/litchi-cfb','crates/litchi-xls'],cwd=ROOT,check=True)
c=json.loads((P/'cleanup.json').read_text());assert c['removed'] and all(not Path(path).exists() for path in c['roots'])
a=json.loads((P/'analysis.json').read_text());assert not a['passed'] and not a['hard_gates']['native'] and a['hard_gates']['repeat'] and a['hard_gates']['bindings']
i=json.loads((P/'audit-final.log').read_text().split(' ',1)[1]);assert i['hard_gates']['bindings'] and not i['hard_gates']['native'] and i['hard_gates']['repeat']
print('PASS exact baseline restoration, archived candidate, rejected disposition and five cleaned roots')
