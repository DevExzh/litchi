"""Freeze successful source, binaries, protocol and independent readers."""
import driver as d
import qualify as q
p=d.P;d.check();source=d.source();assert source==d.read(p/'candidate-source.json')['source']
rows=q.all_rows();assert len(rows)==54
for lane in ['qualification','observer','allocation-preflight']:
 assert {(x['case'],x['leg']) for x in rows if x['lane']==lane}=={(case,leg) for case in d.read(p/'plan.json')['cases'] for leg in ['before','after']}
# Read the event traces as a physical mechanism gate, independent of timing.
mechanism=[]
for case in d.read(p/'plan.json')['cases']:
 pair={x['leg']:x for x in rows if x['lane']=='observer' and x['case']==case}
 before=pair['before']['observer_metrics'];after=pair['after']['observer_metrics']
 expected_gap=512 if case=='cfb-difat' else 0
 assert after['zero_filled_gap_bytes']==expected_gap,(case,after)
 for key in ['write_calls','write_bytes','seek_calls']:assert before[key]==after[key],(case,key)
 assert after['backward_seek_calls']==(1 if case=='cfb-difat' else 0)
 mechanism.append(dict(case=case,before=before,after=after))
d.write(p/'mechanism.json',dict(status='pass',rows=mechanism,scope='Untimed cursor high-water replay; gap bytes are overwritten zero initialization, not total memory copies or physical page faults.'))
for leg in ['before','after']:
 for name,key in [(f'quality-{leg}.json','receipts'),(f'probe-quality-{leg}.json','receipts')]:
  value=d.read(p/name);assert value['status']=='pass'
  for x in value[key]:q.bind(x);assert d.read(x['path'])['exit_code']==0
 for kind in ['native','allocation']:
  b=d.read(p/f'build-{leg}-{kind}.json');q.bind(b['binary']);q.bind(b['freeze']);q.bind(b['receipt']);assert d.read(b['receipt']['path'])['exit_code']==0
  frozen=d.read(b['freeze']['path']);assert frozen['source']==d.read(p/('source-before.json' if leg=='before' else 'candidate-source.json'))['source']
  assert all(d.sha(p/n)==h for n,h in frozen['probe'].items())
assert d.read(p/'reader-tests.json')['status']=='pass';q.bind(d.read(p/'reader-tests.json')['reader'])
files=[]
for file in sorted(p.iterdir()):
 if file.is_file() and file.name not in ['admission.json'] and (file.suffix in ['.py','.json','.patch']):files.append(d.desc(file))
for folder in ['probe','source-before','source-after']:
 files.extend(d.desc(file) for file in sorted((p/folder).rglob('*')) if file.is_file())
d.write(p/'admission.json',dict(status='pass',source=source,files=files,qualified_reports=54))
print('admission PASS',len(files),'bound files')
