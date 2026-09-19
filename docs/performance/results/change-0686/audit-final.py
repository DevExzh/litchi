#!/usr/bin/env python3
"""Final quality/build/corpus/profile bindings, in addition to the measurement audits."""
import hashlib,json,re,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
base=read(P/'baseline.json')
for n,h in base['constraints_sha256'].items():assert sha(ROOT/n)==h,n
quality=read(P/'final-verified/results.json');assert len(quality)==6
for r in quality:
 assert r['exit_code']==0,r['name']
 for n,h in r['source_sha256'].items():assert sha(ROOT/n)==h,n
counts={}
for name in ['tests','facade']:
 matches=re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;', (P/'final-verified'/(name+'.log')).read_text())
 counts[name]={'passed':sum(int(r[0]) for r in matches),'failed':sum(int(r[1]) for r in matches),'ignored':sum(int(r[2]) for r in matches)}
 assert counts[name]['failed']==0 and counts[name]['passed']>0
for phase in ['baseline','candidate']:
 builds=read(P/(phase+'-builds.json'));assert len(builds)==4
 for r in builds:
  assert r['exit_code']==0
  for n,h in r['source_sha256'].items():
   actual=sha(ROOT/n) if phase=='candidate' else hashlib.sha256(subprocess.check_output(['git','show',base['baseline_head']+':'+n],cwd=ROOT)).hexdigest()
   assert actual==h,n
 d=P/'corpus'/phase;m=read(d/'manifest.json')
 assert m['binary_sha256']==next(r['binary_sha256'] for r in builds if r['binary']=='xls-index-probe-0684')
 assert m['corpus_manifest_sha256']==sha(P/'corpus-manifest.json')
 for n,h in m['raw_sha256'].items():assert sha(d/n)==h,n
 for n,h in m['probe_sha256'].items():assert sha(ROOT/n)==h,n
 assert m['source_sha256']==builds[0]['source_sha256']
 for mode in ['owned','file']:
  j=read(d/(mode+'.json'));assert j['query_mismatches']==0 and j['files_seen']==126
  assert j==read(P/'corpus/baseline'/(mode+'.json'))
 for folder in ['native','costs','diagnostics']:
  manifest=read(P/folder/phase/'manifest.json');assert manifest['source_sha256']==builds[0]['source_sha256']
  for r in builds:
   expected=manifest.get('binary_sha256')
   if isinstance(expected,dict):expected=expected.get(r['binary'])
   elif r['binary']!='xls-index-retry-probe-0686':expected=None
   if folder=='native' and r['binary']=='xls-index-retry-probe-0686':expected=manifest['before_binary_sha256' if phase=='baseline' else 'after_binary_sha256']
   if expected:assert expected==r['binary_sha256']
 profiles=read(P/'profiles-manifest.json')
 for n,h in profiles[phase]['files'].items():assert sha(P/n)==h,n
 assert profiles[phase]['diagnostic_manifest_sha256']==sha(P/'diagnostics'/phase/'manifest.json')
 command=read(P/(phase+'-profile-command.json'));assert command['exit_code']==command['report_exit_code']==0
 assert command['binary_sha256']==next(r['binary_sha256'] for r in builds if r['binary']=='xls-index-retry-probe-0686')
 j=read(P/(phase+'-profile.json'));assert j['queries']==10000 and j['warmups']==1 and j['samples']==1 and j['records'][0]['all_queries_agree']
for r in read(P/'evidence/results.json'):assert r['exit_code']==0
print('PASS final quality/build/corpus/profile bindings; test counts:',counts)
(P/'test-counts.json').write_text(json.dumps(counts,indent=2)+'\n')
initial=P/'initial-before-review'
for r in read(initial/'final-verified/results.json'):
 assert r['exit_code']==0
 for n,h in r['source_sha256'].items():
  f=initial/n
  assert sha(f if f.exists() else ROOT/n)==h,n
print('PASS archived pre-review source/check bindings; archived candidate is not the retained implementation')
def untimed(v):
 if isinstance(v,dict):return {k:untimed(x) for k,x in v.items() if 'elapsed' not in k}
 if isinstance(v,list):return [untimed(x) for x in v]
 return v
generated=read(P/'generator-manifest.json');assert generated['identical_two_runs']
for field in ['template','generator_sha256']:
 for n,h in generated[field].items():assert sha(ROOT/n)==h,n
for n,info in generated['files'].items():assert sha(ROOT/n)==info['sha256'] and (ROOT/n).stat().st_size==info['bytes']
for phase in ['baseline','candidate']:
 d=P/'corpus'/phase;m=read(d/'manifest.json')
 for n,h in m['generated_fixtures'].items():assert sha(ROOT/n)==h,n
 for n in [70000,100000]:
  for mode in ['owned','file']:
   filename=f'synthetic-{n}-default-{mode}.json';j=read(d/filename)
   assert j['records'][0]['visit']['actual_callbacks']==n+1
   assert j['records'][0]['visit']['oracle_callbacks']==n+1
   assert untimed(j)==untimed(read(P/'corpus/baseline'/filename))
print('PASS two-run generated fixture hashes and independent full-visitor counts/digests for both sources')
