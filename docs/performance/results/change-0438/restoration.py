#!/usr/bin/env python3
"""Prove the rejected source delta is archived and the checkout is restored."""
from pathlib import Path
import importlib.util,json,hashlib,subprocess
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[3]
spec=importlib.util.spec_from_file_location('restoration_custody',ROOT/'check.py')
c=importlib.util.module_from_spec(spec);spec.loader.exec_module(c)
b=json.loads((ROOT/'before/build.json').read_text())
a=json.loads((ROOT/'after/build.json').read_text())
assert c.sources()==b['source_manifest']
before=json.loads((ROOT/b['source_manifest']['path']).read_text())
after=json.loads((ROOT/a['source_manifest']['path']).read_text())
for retained,source,manifest in [('before-streaming.rs.txt','crates/litchi-odp/src/streaming.rs',before),('after-streaming.rs.txt','crates/litchi-odp/src/streaming.rs',after),('markup_tests.rs.txt','crates/litchi-odp/src/streaming/markup_tests.rs',after)]:
 assert hashlib.sha256((ROOT/'candidate'/retained).read_bytes()).hexdigest()==manifest[source]
for item in json.loads((ROOT/'candidate-artifacts.json').read_text()):
 data=(ROOT/item['path']).read_bytes()
 assert len(data)==item['bytes'] and hashlib.sha256(data).hexdigest()==item['sha256']
subprocess.run(['git','apply','--check',str(ROOT/'candidate/candidate.patch')],cwd=REPO,check=True)
assert hashlib.sha256((REPO/'docs/GOAL.md').read_bytes()).hexdigest()=='bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1'
print('PASS: baseline source manifest restored; candidate source/patch custody and user goal verified')
