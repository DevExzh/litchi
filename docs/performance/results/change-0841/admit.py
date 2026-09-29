"""Freeze qualified sources/readers/receipts before comparative capture."""
import driver as d
import qualify as q
assert d.source()==d.read(d.P/'source-after.json')['source']
assert q.qualification()==d.read(d.P/'qualification.json')
for leg in ['before','after']:
 for name in [f'quality-{leg}.json',f'probe-quality-{leg}.json']:
  value=d.read(d.P/name);assert value['status']=='pass'
  for x in value['receipts']:assert d.read(q.bind(x))['exit_code']==0
 build=d.read(d.P/f'build-{leg}.json');q.bind(build['binary']);q.bind(build['freeze']);assert d.read(q.bind(build['receipt']))['exit_code']==0
 symbols=d.read(d.P/f'symbols-{leg}.json');assert symbols['status']=='pass' and symbols['binary']==build['binary']
 for v in symbols['inventory'].values():q.bind(v)
 for s in symbols['selected']:q.bind(s['assembly'])
assert d.read(d.P/'reader-tests.json')['status']=='pass'
files=[d.desc(p) for p in sorted(d.P.iterdir()) if p.is_file() and p.suffix in ['.py','.json','.patch']]
files += [d.desc(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()]
d.write(d.P/'admission.json',dict(status='pass',files=files))
print('formal admission PASS',len(files),'files')
