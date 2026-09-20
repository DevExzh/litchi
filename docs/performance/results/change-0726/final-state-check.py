#!/usr/bin/env python3
"""Read-only source, constraints, disposition and scratch-cleanup verification."""
import hashlib,json
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
for rel,digest in json.loads((P/'constraints.json').read_text()).items():assert sha(ROOT/rel)==digest,rel
r=json.loads((P/'disposition.json').read_text());assert r['exact']
actual={str(f.relative_to(ROOT)):sha(f) for name in ['litchi-cfb','litchi-xls'] for f in (ROOT/'crates'/name).rglob('*.rs')}
assert actual==r['final_rust_source_sha256']
for rel,digest in r['candidate_source_sha256'].items():assert sha(P/'candidate-source'/rel)==digest
c=json.loads((P/'cleanup.json').read_text());assert c['removed'] and all(not Path(path).exists() for path in c['roots'])
a=json.loads((P/'analysis.json').read_text());assert r['disposition']==('retained' if a['passed'] else 'rejected')
i=json.loads((P/'audit-final.log').read_text().split(' ',1)[1]);assert i['hard_gates']['bindings']
assert all(i['hard_gates'].values())==a['passed']
t=json.loads((P/'terminal-replay.json').read_text());assert len(t['commands'])==4 and all(x['exit_code']==x['expected_exit_code'] for x in t['commands'])
print('PASS exact final source, archived candidate, disposition, constraints, offline replay and three cleaned roots')
