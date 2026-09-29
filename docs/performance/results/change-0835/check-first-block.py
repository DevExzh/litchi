"""Read-only check of the completed first formal block while capture continues."""
import reader
import driver as d
binding=reader._prepare_binding(d.P)
pids=set();environment=None;rows=[]
for i,planned in enumerate(d.read(d.P/'measurement-plan.json')['rows'][:12]):
 report=d.P/f'native-{i:03}.json'
 checked=reader._validate_single(report,planned['case'],(planned['cache_state'],),30,3,binding,environment,pids)
 environment=environment or checked['environment']
 rows.append(d.desc(report))
assert len(pids)==360
d.write(d.P/'first-block-validation.json',dict(status='pass',reports=rows,sample_count=360,reader=d.desc(d.P/'reader.py')))
print('first complete formal block PASS: 12 reports / 360 samples')
