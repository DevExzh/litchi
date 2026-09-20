#!/usr/bin/env python3
"""Retain only a passing candidate; otherwise restore the one archived path."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
a=json.loads((P/'analysis.json').read_text());rel='crates/litchi-xls/src/workbook/source.rs'
assert (ROOT/rel).read_bytes()==(P/'candidate-source'/rel).read_bytes()
raw=(P/'audit-final.log').read_text();assert raw.startswith('PASS ' if a['passed'] else 'REJECTED ')
expected=json.loads((P/('candidate-builds.json' if a['passed'] else 'baseline-builds.json')).read_text())[0]['source_sha256']
if not a['passed']:
 (ROOT/rel).write_bytes(subprocess.check_output(['git','show','3ad29e42da:'+rel],cwd=ROOT))
actual={str(f.relative_to(ROOT)):sha(f) for name in ['litchi-cfb','litchi-xls'] for f in (ROOT/'crates'/name).rglob('*.rs')}
assert actual==expected
result=dict(disposition='retained' if a['passed'] else 'rejected',baseline_revision=subprocess.check_output(['git','rev-parse','3ad29e42da'],cwd=ROOT,text=True).strip(),candidate_source_sha256=json.loads((P/'candidate-source.json').read_text()),final_rust_source_sha256=actual,exact=True)
(P/'disposition.json').write_text(json.dumps(result,indent=2)+'\n');print('PASS',result['disposition'],'exact',len(actual),'Rust source files')
