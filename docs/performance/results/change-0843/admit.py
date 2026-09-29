"""Freeze complete successful qualification and reader/observer evidence."""
import driver as d
import qualify as q
import archive_check_v2 as ac
import observer_check as oc

d.check();assert d.source()==d.read(d.P/'source.json')['source']
assert d.sha(d.ROOT/'Cargo.lock')==d.sha(d.P/'root-Cargo.lock')
assert q.qualification()==d.read(d.P/'qualification.json')
assert ac.verify()==d.read(d.P/'archive-check-v2.json')
assert oc.verify()==d.read(d.P/'observer-preflight-validation.json')
for name in ['reader-tests','heap-reader-tests']:
 value=d.read(d.P/(name+'.json'));assert value['status']=='pass'
 for key in ['reader','test_source','receipt']:q.bind(value[key])
 assert d.read(value['receipt']['path'])['exit_code']==0
build=d.read(d.P/'build.json');q.bind(build['binary']);q.bind(build['freeze'])
assert d.read(d.P/'symbols.json')['status']=='pass' and d.read(d.P/'symbols.json')['binary']==build['binary']
assert d.read(d.P/'production-gate-reuse.json')['status']=='pass'
for folder in (d.P/'commands').iterdir():
 end=d.read(folder/'receipt.json');q.bind(end['log'])
 assert end['exit_code']==(101 if folder.name in ['probe-clippy','probe-v2-clippy'] else 0)
assert len(list((d.P/'runs').iterdir()))==8
files=[d.desc(p) for p in sorted(d.P.rglob('*')) if p.is_file()]
d.write(d.P/'admission.json',dict(status='pass',files=files,note='Single current-source binary;8archive qualifications and2separate observer preflights completed. No formalnative/profile/heapcapture exists yet.'))
print('Admission PASS;',len(files),'immutable files')
