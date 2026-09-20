#!/usr/bin/env python3
"""Restore only the three archived candidate paths after rejection."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
paths=['crates/litchi-xls/src/workbook/query_cache.rs','crates/litchi-xls/src/workbook/source.rs','crates/litchi-xls/tests/xls_query_chain_checkpoint.rs']
def sha(b):return hashlib.sha256(b).hexdigest()
assert not json.loads((P/'analysis.json').read_text())['passed']
for rel in paths:assert (ROOT/rel).read_bytes()==(P/'candidate-source'/rel).read_bytes(),rel
rows=[]
for rel in paths:
 p=ROOT/rel;candidate=sha(p.read_bytes())
 r=subprocess.run(['git','show','45cb480eaa:'+rel],cwd=ROOT,capture_output=True)
 if r.returncode==0:p.write_bytes(r.stdout);baseline=sha(r.stdout)
 else:
  assert rel==paths[-1];p.unlink();baseline=None
 rows.append(dict(path=rel,candidate_sha256=candidate,baseline_sha256=baseline,removed=baseline is None))
expected=json.loads((P/'baseline-builds.json').read_text())[0]['source_sha256']
actual={str(p.relative_to(ROOT)):sha(p.read_bytes()) for name in ('litchi-cfb','litchi-xls') for p in (ROOT/'crates'/name).rglob('*.rs')}
assert actual==expected
subprocess.run(['git','diff','--exit-code','45cb480eaa','--','crates/litchi-cfb','crates/litchi-xls'],cwd=ROOT,check=True)
(P/'restoration.json').write_text(json.dumps(dict(disposition='rejected',baseline_revision=subprocess.check_output(['git','rev-parse','45cb480eaa'],cwd=ROOT,text=True).strip(),paths=rows,baseline_rust_source_count=len(actual),baseline_rust_source_sha256=actual,exact=True),indent=2)+'\n')
print('PASS baseline restored; exact',len(actual),'Rust source files; candidate preserved in packet')
