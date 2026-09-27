"""Replay quality receipts, paired samples, and the expanded-name refusal oracle."""
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
 initial=read(P/'quality-0/manifest.json');assert initial['rows'][-1]['exit']!=0
 assert sha(P/initial['source_file'])==initial['source_sha256']
 for row in initial['rows']:assert sha(P/row['log'])==row['sha256']
 final=read(P/'final-source.json');tested=read(P/q['source_file'])
 assert all(tested[n]==h for n,h in final['files'].items())
 d=P/'measure-0'
 for row in read(d/'build.json'):assert row['exit']==0 and sha(d/row['log'])==row['sha256']
 binaries=read(d/'binaries.json');probe=read(d/'probe-inputs.json')
 for leg in ['before','after']:
  assert sha(d/leg/'src/main.rs')==probe['main.rs']
  assert sha(d/leg/'Cargo.lock')==binaries[leg]['lock_sha256']
 assert binaries['before']['lock_sha256']==binaries['after']['lock_sha256']
 a=analyze.analyze();assert a==read(P/'analysis.json')
 rows=read(d/'runs.json');assert len(rows)==110 and len(a['cases'])==18
 assert [r['kind'] for r in rows[:2]]==['differential']*2
 for offset in range(2,len(rows),6):
  group=rows[offset:offset+6];assert len({r['case'] for r in group})==1
  assert [r['leg'] for r in group]==['before','after','after','before','before','after']
 before=read(d/'differential-before.json');after=read(d/'differential-after.json')
 assert not before['alias_oracle_failures'] and not after['alias_oracle_failures']
 for spelling in ['aliased','direct']:
  assert after['results']['alias_pair::invalid-expanded-duplicate::'+spelling]=='invalid-control:codec_rejected=true:stream_rejected=true'
 seal=P/'seal.json'
 if seal.exists():
  files={str(f.relative_to(P)):sha(f) for f in sorted(P.rglob('*')) if f.is_file() and f!=seal}
  assert files==read(seal)['files'],'packet seal mismatch'
 return {'gates':len(q['rows']),'tests':counts,'cases':len(a['cases']),'differential_changes':len(a['differential']['changes'])}
if __name__=='__main__':print(json.dumps(validate(),indent=2))
