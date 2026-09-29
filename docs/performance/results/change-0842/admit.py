"""Freeze all completed input/equivalence/gate evidence before native capture."""
import driver as d
import qualify as q
import codec_check_v2 as c
import json

d.check();assert d.source()==d.read(d.P/'source-final.json')['source']
assert q.qualification()==d.read(d.P/'qualification.json')
assert c.verify()==d.read(d.P/'codecs-validation-v2.json')
assert d.read(d.P/'quality-after.json')['status']=='pass'
for label in ['before-opc-check','before-opc-test','before-opc-clippy','before-opc-doc','before-fmt','before-boundaries']:
 assert d.read(d.P/'commands'/label/'receipt.json')['exit_code']==0
for leg in ['before','after']:
 build=d.read(d.P/f'build-{leg}.json');q.bind(build['binary']);q.bind(build['codec']);q.bind(build['freeze'])
 for n,h in d.read(build['freeze']['path'])['probe'].items():
  path=d.P/'lock-snapshots'/f'{leg}.lock' if n=='probe/Cargo.lock' else d.P/n
  assert d.sha(path)==h
checks=d.read(d.P/'reader-tests.json');assert checks['status']=='pass' and len(checks['checks'])==19
q.bind(checks['reader']);q.bind(checks['tests']);q.bind(checks['receipt'])
assert d.read(d.P/'feature-diff.json')['status']=='pass'
assert d.read(d.P/'nested-lock-checks.json')['status']=='pass'
assert d.read(d.P/'tool-checks.json')['status']=='pass'
for item in d.read(d.P/'tool-checks.json')['receipts']:
 assert d.read(q.bind(item))['exit_code']==0
# A writer left running here would invalidate the freeze, so every started
# command must already have an exit receipt and a stable, hash-bound log.
for folder in (d.P/'commands').iterdir():
 result=d.read(folder/'receipt.json');q.bind(result['log'])
files=[d.desc(p) for p in sorted(d.P.rglob('*')) if p.is_file()]
d.write(d.P/'admission.json',dict(status='pass',files=files,source_after=d.desc(d.P/'source-after.json'),note='Includes both stage-specific locks, all final qualifications, codec artifacts, exact feature diff, preflight failures and all gates. No formal native capture exists yet.'))
assert not any(p.name.startswith('native-') for p in (d.P/'runs').iterdir())
print('Native admission PASS;',len(files),'frozen files')
