"""Replay quality receipts, paired samples, and the expanded-name refusal oracle."""
from pathlib import Path
import hashlib,json,re,runpy
import analyze
import triage
P=Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def checked_quality(manifest_path,expected_files=None):
 manifest=read(manifest_path);assert len(manifest['rows'])==9
 source=read(P/manifest['source_file']);assert sha(P/manifest['source_file'])==manifest['source_sha256']
 if expected_files is not None:assert all(source[name]==digest for name,digest in expected_files.items())
 counts={'passed':0,'ignored':0,'suites':0}
 for row in manifest['rows']:
  assert row['exit']==0 and sha(P/row['log'])==row['sha256']
  matches=re.findall(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored',(P/row['log']).read_text())
  assert all(m[0]=='ok' and m[2]=='0' for m in matches)
  counts['passed']+=sum(int(m[1]) for m in matches);counts['ignored']+=sum(int(m[3]) for m in matches);counts['suites']+=len(matches)
 return manifest,counts
def validate_triage(a):
 t=read(P/'triage.json');changes=a['differential']['changes'];assert t['comparisons']==a['differential']['comparisons'] and t['changes']==changes
 expected={};groups={};before_outcomes={'ok':0,'err':0}
 for key,item in changes.items():
  parts=key.split('::');groups[parts[0]]=groups.get(parts[0],0)+1
  before_outcomes['ok' if item['before'].startswith('ok:') else 'err']+=1
  if parts[0]=='alias_pair':
   assert key=='alias_pair::invalid-expanded-duplicate::aliased' and item['after']=='invalid-control:codec_rejected=true:stream_rejected=true';continue
  assert parts[0] in ['aliasing','generated'] and len(parts)>=3 and parts[2].startswith('mce_')
  expected.setdefault('::'.join(parts[:2]),[]).append(item)
 assert t['groups']==groups and t['before_outcomes']==before_outcomes and set(t['documents'])==set(expected)
 generator=t['generator'];g=P/'generator-1.rs';log=P/'generator-1-build.log'
 assert generator['binary_removed'] is True and sha(g)==generator['source_sha256'] and sha(log)==generator['log_sha256']
 input_dirs=set()
 for doc,records in t['documents'].items():
  path=(P/records['input']).resolve();input_dirs.add(path.parent)
  assert path.parent.name=='changed-inputs-1' and path.is_file() and sha(path)==records['sha256'] and records['tokenizer_error'] is None
  assert triage.witnesses(path.read_bytes())=={field:records[field] for field in ['duplicates','unbound_attributes','tokenizer_error']}
  assert records['duplicates']
  assert all(item['after']=='err:non-conformant markup compatibility XML: duplicate attribute' for item in expected[doc])
 assert input_dirs=={(P/'changed-inputs-1').resolve()}
 assert {p.name for p in input_dirs.pop().iterdir() if p.is_file()}=={Path(item['input']).name for item in t['documents'].values()}
 return len(t['documents'])
def validate_allocations(a,cleanup_targets):
 d=P/'measure-1';out=P/'allocations-0';assert read(out/'complete.json')['source_unchanged'] is True
 bins=read(d/'binaries.json');runs=read(out/'runs.json');expected_cases={'mce_benign_worksheet','mce_benign_document','mce_prefixed_32'}
 assert len(runs)==6 and {(row['case'],row['leg']) for row in runs}=={(case,leg) for case in expected_cases for leg in ['before','after']}
 for row in runs:
  assert row['exit']==0 and row['decode_exit']==0 and row['leg'] in ['before','after'] and row['case'] in expected_cases
  for field in ['stdout','stderr','report','capture','summary','histogram']:
   path=out/row[field];assert not Path(row[field]).is_absolute() and path.parent==out and path.is_file() and sha(path)==row[field+'_sha256']
  command=row['command'];adversarial=command.index('adversarial');assert adversarial>0 and command[adversarial-1]==bins[row['leg']]['path']
  report=read(out/row['report']);expected=a['cases'][row['case']]
  assert report['case']==row['case'] and report['samples']==1 and report['warmup']==0 and len(report['elapsed_ns'])==1 and report['p50_ns']==report['elapsed_ns'][0]
  assert report['input_bytes']==expected['input_bytes'] and report['input_sha256']==expected['input_sha256'] and report['outcomes']==[expected['outcome']]
  histogram=[]
  for line in (out/row['histogram']).read_text().splitlines():
   fields=line.split();assert len(fields)==2 and all(re.fullmatch(r'[0-9]+',field) for field in fields)
   size,count=map(int,fields);assert size>0 and count>0;histogram.append((size,count))
  assert histogram
  calls=sum(count for _,count in histogram);allocated=sum(size*count for size,count in histogram)
  assert row['allocation_calls']==calls and row['allocated_bytes']==allocated
  match=re.search(r'calls to allocation functions: ([0-9]+)',(out/row['summary']).read_text());assert match and int(match[1])==calls
  if cleanup_targets is None:
   binary=Path(bins[row['leg']]['path']);assert binary.is_file() and sha(binary)==bins[row['leg']]['sha256']
 return len(runs)
def validate_cleanup():
 receipt=P/'cleanup.json'
 if not receipt.exists():return None
 cleanup=read(receipt);assert cleanup['executables_verified_before_removal'] is True and cleanup['source_unchanged'] is True and cleanup['fixtures_unchanged'] is True
 expected={Path('/home/zhuhe/code/litchi-target-0776-quality'),Path('/home/zhuhe/code/litchi-target-0776-release-before'),Path('/home/zhuhe/code/litchi-target-0776-release-after')}
 rows=cleanup['targets'];assert len(rows)==3 and {Path(row['path']) for row in rows}==expected
 for row in rows:
  path=Path(row['path']);assert path.parent==Path('/home/zhuhe/code') and path.name.startswith('litchi-target-0776-') and row['removed'] is True and isinstance(row['bytes'],int) and row['bytes']>=0 and not path.exists() and not path.is_symlink()
 return len(rows)
def validate():
 q,counts=checked_quality(P/'quality.json')
 first=read(P/'first-source.json');_,intermediate_counts=checked_quality(P/'quality-1/manifest.json',first['files'])
 initial=read(P/'quality-0/manifest.json');assert initial['rows'][-1]['exit']!=0
 assert sha(P/initial['source_file'])==initial['source_sha256']
 for row in initial['rows']:assert sha(P/row['log'])==row['sha256']
 final=read(P/'final-source.json');tested=read(P/q['source_file'])
 assert all(tested[n]==h for n,h in final['files'].items())
 d=P/'measure-1'
 for row in read(d/'build.json'):assert row['exit']==0 and sha(d/row['log'])==row['sha256']
 binaries=read(d/'binaries.json');probe=read(d/'probe-inputs.json')
 source_maps={leg:read(d/f'source-{leg}.json') for leg in ['before','after']};workspace_lock=sha(P/'workspace-Cargo.lock')
 assert all(source_maps[leg]['Cargo.lock']==workspace_lock for leg in ['before','after'])
 assert all(source_maps['after'][name]==digest for name,digest in tested.items())
 first_measured=read(P/'measure-0/source-after.json');assert all(first_measured[name]==digest for name,digest in first['files'].items())
 assert sha(P/'probe-src/main.rs')==probe['main.rs'] and sha(P/'probe-src/Cargo.toml.template')==probe['Cargo.toml.template']
 for leg in ['before','after']:
  assert sha(d/leg/'src/main.rs')==probe['main.rs']
  assert sha(d/leg/'Cargo.lock')==binaries[leg]['lock_sha256']
 assert binaries['before']['lock_sha256']==binaries['after']['lock_sha256']
 a=analyze.analyze();assert a==read(P/'analysis.json')
 replayed=runpy.run_path(str(P/'analyze-0.py'),run_name='validate-analyze-0')['analyze']();assert replayed==read(P/'analysis-0.json')
 rows=read(d/'runs.json');assert len(rows)==110 and len(a['cases'])==18
 assert [r['kind'] for r in rows[:2]]==['differential']*2
 for offset in range(2,len(rows),6):
  group=rows[offset:offset+6];assert len({r['case'] for r in group})==1
  assert [r['leg'] for r in group]==['before','after','after','before','before','after']
 before=read(d/'differential-before.json');after=read(d/'differential-after.json')
 assert not before['alias_oracle_failures'] and not after['alias_oracle_failures']
 for spelling in ['aliased','direct']:
  assert after['results']['alias_pair::invalid-expanded-duplicate::'+spelling]=='invalid-control:codec_rejected=true:stream_rejected=true'
 triage_documents=validate_triage(a)
 cleanup_targets=validate_cleanup()
 allocation_runs=validate_allocations(a,cleanup_targets)
 seal=P/'seal.json'
 if seal.exists():
  files={str(f.relative_to(P)):sha(f) for f in sorted(P.rglob('*')) if f.is_file() and f!=seal}
  assert files==read(seal)['files'],'packet seal mismatch'
 return {'gates':len(q['rows']),'tests':counts,'intermediate_gates':9,'intermediate_tests':intermediate_counts,'cases':len(a['cases']),'differential_changes':len(a['differential']['changes']),'triage_documents':triage_documents,'allocation_runs':allocation_runs,'cleanup_targets':cleanup_targets}
if __name__=='__main__':print(json.dumps(validate(),indent=2))
