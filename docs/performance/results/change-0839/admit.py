"""Bind final compared sources, binaries, quality, readers and frozen policy."""
import driver as d
source=d.source();assert source==d.read(d.P/'candidate-source.json')['source']
for leg in ['before','after']:
 freeze=d.read(d.P/f'freeze-{leg}.json');quality=d.read(d.P/f'quality-{leg}.json')
 assert freeze['source']==quality['source']
 assert freeze['probe']==quality['probe']=={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
 for x in quality['commands']:assert d.read(x['path'])['exit_code']==0
 allocator_leg='before-v2' if leg=='before' else 'after'
 aq=d.read(d.P/f'allocator-quality-{allocator_leg}.json');af=d.read(d.P/f'freeze-{leg}-allocator.json')
 assert aq['source']==af['source']==freeze['source']
 assert aq['inputs']==af['inputs']=={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'allocator').rglob('*')) if p.is_file()}
 for x in aq['receipts']:assert d.read(x['path'])['exit_code']==0
 for suffix in ['', '-observer','-allocator']:
  b=d.read(d.P/f'build-{leg}{suffix}.json');assert d.desc(b['binary']['path'])==b['binary'];assert d.read(b['receipt']['path'])['exit_code']==0
assert d.read(d.P/'qualification-replay-both.json')['status']=='pass'
files=[d.P/n for n in ['plan.json','plan-amendment.json','driver.py','build.py','capture.py','qualify.py','readers.py','analyze.py','trace.py','reader-tests.json','launcher-quality.json','candidate-source.json','candidate.patch','qualification-replay-both.json']]
files += list(d.P.glob('build-*.json'))+list(d.P.glob('freeze-*.json'))+list(d.P.glob('quality-*.json'))+list(d.P.glob('allocator-quality-*.json'))
files += [p for folder in ['probe','allocator'] for p in sorted((d.P/folder).rglob('*')) if p.is_file()]
d.write(d.P/'admission.json',dict(status='pass',source=source,files=[d.desc(p) for p in sorted(set(files))]))
print('admission PASS')
