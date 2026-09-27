"""Offline receipt, source, statistics, and packet-inventory validation."""
from pathlib import Path
import hashlib,json,re
import analyze
P=Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def validate():
 q=read(P/'quality.json');assert len(q['rows'])==9
 assert sha(P/q['source_file'])==q['source_sha256']
 counts={'passed':0,'ignored':0,'suites':0}
 for row in q['rows']:
  assert row['exit']==0 and sha(P/row['log'])==row['sha256']
  matches=re.findall(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored',(P/row['log']).read_text())
  assert all(m[0]=='ok' and m[2]=='0' for m in matches)
  counts['passed']+=sum(int(m[1]) for m in matches);counts['ignored']+=sum(int(m[3]) for m in matches);counts['suites']+=len(matches)
 for folder in ['measure-0','measure-1']:
  for row in read(P/folder/'build.json'):
   assert row['exit']==0 and sha(P/folder/row['log'])==row['sha256']
  runs=read(P/folder/'runs.json');assert len(runs)==1 and runs[0]['kind']=='differential' and runs[0]['exit']!=0
  for field in ['report','log']:assert sha(P/folder/runs[0][field])==runs[0][field+'_sha256']
 d=P/'measure-2'
 for row in read(d/'build.json'):
  assert row['exit']==0 and sha(d/row['log'])==row['sha256']
 binaries=read(d/'binaries.json');probe=read(d/'probe-inputs.json')
 for leg in ['before','after']:
  assert sha(d/leg/'src/main.rs')==probe['main.rs']
  assert sha(d/leg/'Cargo.lock')==binaries[leg]['lock_sha256']
 assert binaries['before']['lock_sha256']==binaries['after']['lock_sha256']
 c=read(P/'comment-followup.json');assert c['only_rust_comments_changed']
 assert sha(P/'comment-only-followup.patch')==c['patch_sha256']
 for row in c['checks']:assert row['exit']==0 and sha(P/row['log'])==row['sha256']
 a=analyze.analyze();assert a==read(P/'analysis.json')
 assert len(read(d/'runs.json'))==110 and len(a['cases'])==18
 rows=read(d/'runs.json');assert [r['kind'] for r in rows[:2]]==['differential']*2
 for offset in range(2,len(rows),6):
  group=rows[offset:offset+6];assert len({r['case'] for r in group})==1
  assert [r['leg'] for r in group]==['before','after','after','before','before','after']
 before=read(d/'differential-before.json');after=read(d/'differential-after.json')
 assert not before['alias_oracle_failures'] and not after['alias_oracle_failures']
 pre=P/'preid-0';reference=read(pre/'differential.json')
 assert reference['results']==after['results'] and not reference['alias_oracle_failures']
 assert sha(pre/'src/main.rs')==probe['main.rs'] and sha(pre/'Cargo.lock')==binaries['before']['lock_sha256']
 for row in read(pre/'runs.json'):assert row['exit']==0 and sha(pre/row['log'])==row['sha256']
 assert sha(pre/'differential.json')==read(pre/'complete.json')['report_sha256']
 triage=read(P/'alias-triage.json');assert not triage['unmatched'] and not triage['all_preid_differences']
 assert triage['changed_comparisons']==triage['matching_preid_exactly']==len(a['differential']['changes'])==1037
 assert triage['preid_report_sha256']==sha(pre/'differential.json')
 seal=P/'seal.json'
 if seal.exists():
  files={str(f.relative_to(P)):sha(f) for f in sorted(P.rglob('*')) if f.is_file() and f!=seal}
  assert files==read(seal)['files'],'packet seal mismatch'
 return {'gates':len(q['rows']),'tests':counts,'cases':len(a['cases']),'differential_changes':len(a['differential']['changes'])}
if __name__=='__main__':print(json.dumps(validate(),indent=2))
