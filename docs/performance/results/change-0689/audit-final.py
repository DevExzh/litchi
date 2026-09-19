#!/usr/bin/env python3
"""Bind final checks, builds, corpus and profiles to the measured sources."""
import hashlib,json,re,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
base=read(P/'baseline.json')
for n,h in base['constraints_sha256'].items():assert sha(ROOT/n)==h,n
quality=read(P/'final-verified/results.json');assert len(quality)==6
for r in [*quality,read(P/'consumer-tests.json')]:
 assert r['exit_code']==0
 for n,h in r['source_sha256'].items():assert sha(ROOT/n)==h,n
counts={}
for name,path in [('owners',P/'final-verified/tests.log'),('facade',P/'final-verified/facade.log'),('doc-ppt',P/'consumer-tests.log')]:
 matches=re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;',path.read_text());assert matches
 counts[name]={k:sum(int(r[i]) for r in matches) for i,k in enumerate(['passed','failed','ignored'])};assert counts[name]['failed']==0
for phase in ['baseline','candidate']:
 builds=read(P/(phase+'-builds.json'));assert len(builds)==5 and all(b['exit_code']==0 for b in builds)
 for r in builds:
  for n,h in r['source_sha256'].items():
   actual=sha(ROOT/n) if phase=='candidate' else hashlib.sha256(subprocess.check_output(['git','show',base['baseline_head']+':'+n],cwd=ROOT)).hexdigest()
   assert actual==h,n
 for folder in ['native','costs','corpus','repeat','diagnostics']:
  m=read(P/folder/phase/'manifest.json');assert m['source_sha256']==builds[0]['source_sha256']
  binaries={b['binary']:b['binary_sha256'] for b in builds}
  if folder=='native':assert m['before_binary_sha256' if phase=='baseline' else 'after_binary_sha256']==binaries['xls-index-retry-probe-0686']
  elif folder=='costs':assert all(h==binaries[n] for n,h in m['binary_sha256'].items())
  elif folder=='repeat':assert m['binary_sha256'][phase]==binaries['xls0684-repeat']
  elif folder=='diagnostics':assert m['binary_sha256']==binaries['xls0684-repeat']
  else:
   assert m['binary_sha256']==binaries['xls-index-probe-0684'] and m['corpus_manifest_sha256']==sha(P/'corpus-manifest.json')
   for n,h in m['raw_sha256'].items():assert sha(P/folder/phase/n)==h,n
   for n,h in m['probe_sha256'].items():assert sha(ROOT/n)==h,n
   for mode in ['owned','file']:
    j=read(P/folder/phase/(mode+'.json'));assert j['files_seen']==126 and j['query_mismatches']==0
    assert j==read(P/'corpus/baseline'/(mode+'.json'))
 command=read(P/(phase+'-profile-command.json'));assert command['exit_code']==command['report_exit_code']==0
 assert command['binary_sha256']==binaries['xls0684-repeat'] and command['source_sha256']==builds[0]['source_sha256']
 assert 'repeats=2000000\tfound=2000000\t' in (P/(phase+'-profile.tsv')).read_text()
def untimed(v):
 if isinstance(v,dict):return {k:untimed(x) for k,x in v.items() if 'elapsed' not in k}
 if isinstance(v,list):return [untimed(x) for x in v]
 return v
for phase in ['baseline','candidate']:
 for mode in ['owned','file']:
  name='synthetic-70000-default-'+mode+'.json';j=read(P/'corpus'/phase/name)
  assert j['records'][0]['visit']['actual_callbacks']==j['records'][0]['visit']['oracle_callbacks']==70001
  assert untimed(j)==untimed(read(P/'corpus/baseline'/name))
evidence=read(P/'evidence/results.json')
assert {r['name'] for r in evidence}=={'crate-boundaries','claims','claims-structural','report','coverage','non-iwork'} and len(evidence)==6
for r in evidence:assert r['exit_code']==0
initial=P/'initial-checks'
checks=read(initial/'final-verified/results.json')
assert [(r['name'],r['exit_code']) for r in checks]==[('fmt',0),('check',0),('clippy',0),('tests',101)]
for r in checks:
 for n,h in r['source_sha256'].items():assert sha(initial/n if (initial/n).exists() else ROOT/n)==h,n
assert 'left: 544' in (initial/'final-verified/tests.log').read_text()
assert 'right: 0' in (initial/'final-verified/tests.log').read_text()
print('PASS archived initial lifecycle-test failure and differing test-source bindings.')
(P/'test-counts.json').write_text(json.dumps(counts,indent=2)+'\n')
print('PASS quality/build/measurement/corpus/profile/constraint bindings; tests:',counts)
for phase in ['baseline','candidate']:
 manifest=read(P/(phase+'-assembly-manifest.json'))
 assert manifest['binary_sha256']==read(P/(phase+'-builds.json'))[0]['binary_sha256']
 assert sha(P/manifest['section_file'])==manifest['section_sha256']
 for row in manifest['symbols']:
  if row['standalone_symbol']:assert sha(P/row['assembly_file'])==row['assembly_sha256']
print('PASS matched assembly bindings.')
for phase,files in read(P/'profiles-manifest.json').items():
 for name,digest in files.items():assert sha(P/name)==digest,name
doc_checks=read(P/'evidence/final-doc-results.json')
assert {r['name'] for r in doc_checks}=={'report','coverage','non-iwork'} and len(doc_checks)==3
for row in doc_checks:assert row['exit_code']==0
summary=read(P/'validation-summary.json')
assert sum(c['passed'] for c in counts.values())==summary['rust_tests_passed']
assert sum(c['ignored'] for c in counts.values())==summary['rust_tests_ignored']==27
for n,h in summary['source_sha256'].items():assert sha(ROOT/n)==h,n
print('PASS final profile hashes, documentation gates and validation summary.')
