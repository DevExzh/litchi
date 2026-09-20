#!/usr/bin/env python3
"""Check exact source contrasts and restoration independently of the build driver."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];OLD=P.parent/'change-0723'
def sha(b):return hashlib.sha256(b).hexdigest()
def data(v,n):return (P/'sources'/v/'crates/litchi-xls/src/workbook'/n).read_bytes()
def unallow(b):return b.replace(b'    #[allow(dead_code)] // Diagnostic ablation omits checkpoint construction.\n',b'').replace(b'    #[allow(dead_code)] // Diagnostic layout-only field.\n',b'')
for name in ('source.rs','query_cache.rs'):
 rel='crates/litchi-xls/src/workbook/'+name
 assert data('baseline',name)==subprocess.check_output(['git','show','0d943df447:'+rel],cwd=ROOT)
 assert data('full',name)==(OLD/'candidate-source'/rel).read_bytes()
 assert (ROOT/rel).read_bytes()==data('baseline',name)
assert data('layout','source.rs')==data('baseline','source.rs')
for v in ('layout','selection'):assert unallow(data(v,'query_cache.rs'))==data('full','query_cache.rs')
base=data('baseline','source.rs');sel=data('selection','source.rs');full=data('full','source.rs')
start=b'fn replay_indexed_cell(';end=b'\n/// Reads and decodes one indexed occurrence.'
def split(b):a=b.index(start);z=b.index(end,a);return b[:a],b[a:z],b[z:]
bp,bm,bs=split(base);sp,sm,ss=split(sel);fp,fm,fs=split(full)
assert bp==sp and bs==ss and sm==fm
baseline=json.loads((P/'source-baseline.json').read_text())
for variant in ('baseline','layout','selection','full'):
 m=json.loads((P/('source-'+variant+'.json')).read_text());assert m.keys()==baseline.keys()
 changed={r for r in baseline if baseline[r]!=m[r]}
 expected=set() if variant=='baseline' else {'crates/litchi-xls/src/workbook/query_cache.rs'}
 if variant in ('selection','full'):expected.add('crates/litchi-xls/src/workbook/source.rs')
 assert changed==expected
 for rel in changed:assert sha((P/'sources'/variant/rel).read_bytes())==m[rel]
assert all(sha((ROOT/r).read_bytes())==h for r,h in baseline.items())
print('PASS exact four variant contrasts, unchanged CFB and baseline restoration')
