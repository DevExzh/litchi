"""Freeze both successful legs and exact artifact parity before native timing."""
import driver as d
import qualify as q
import observer_check as oc
import profile as ip
import json

original=measurement=d.read(d.P/'plan.json')
assert len(original['qualification'])+len(original['native'])+len(original['heap'])==original['formal_reports']==512
assert sum(row['samples'] for lane in ['qualification','native','heap'] for row in original[lane])==original['formal_samples']==211328
for case in measurement['cases']:
 first=[next(x for x in measurement['native'] if x['block']==b and all(x[k]==v for k,v in case.items()))['leg'] for b in range(6)]
 assert first.count('before')==first.count('after')==3

provenance=d.read(d.P/'harness-provenance.json');assert provenance['status']=='pass'
for item in provenance['files']:
 q.bind(item['current']);q.bind(item['previous']);assert item['current']['sha256']==item['previous']['sha256']
d.check();assert d.source()==d.read(d.P/'source-after.json')['source']
assert d.sha(d.ROOT/'Cargo.lock')==d.sha(d.P/'root-Cargo.lock')
assert q.qualification()==d.read(d.P/'qualification.json')
for leg in ['before','after']:
 capture=d.read(d.P/f'capture-qualification-{leg}.json')
 assert capture['status']=='pass' and capture['plan']==d.desc(d.P/'plan.json')
 assert capture['receipts']==[d.desc(q.folder_for('qualification',row)/'receipt.json') for row in original['qualification'] if row['leg']==leg]
 assert q.qualification(leg)==d.read(d.P/f'qualification-{leg}.json')
 assert oc.verify(leg)==d.read(d.P/f'observer-preflight-{leg}.json')
 b=d.read(d.P/f'build-{leg}.json');q.bind(b['binary']);q.bind(b['freeze'])
 assert d.read(d.P/f'symbols-{leg}.json')['binary']==b['binary']
assert d.read(d.P/'freeze-before.json')['probe']==d.read(d.P/'freeze-after.json')['probe']
a=d.read(d.P/'source-before.json')['source'];b=d.read(d.P/'source-after.json')['source']
assert {k for k in set(a)|set(b) if a.get(k)!=b.get(k)}=={'crates/soapberry-zip/src/office.rs','crates/litchi-opc/src/source_backed.rs'}
assert d.read(d.P/'quality-after.json')['status']=='pass'
for name in ['reader-tests','heap-reader-tests']:
 value=d.read(d.P/(name+'.json'));assert value['status']=='pass'
 for key in ['reader','test_source','receipt']:q.bind(value[key])
 assert d.read(value['receipt']['path'])['exit_code']==0
for folder in (d.P/'commands').iterdir():
 end=d.read(folder/'receipt.json');q.bind(end['log']);assert end['exit_code']==0,(folder.name,end['exit_code'])
assert len(list((d.P/'runs').iterdir()))==64
side=ip.plan();assert side['reports']==side['formal_samples']==16 and not (d.P/'profile-runs').exists()
pre=d.read(d.P/'profile-preflight.json');assert pre['status']=='pass' and len(pre['qualification_reports'])==8
assert pre['plan']==d.desc(d.P/'profile-plan.json') and pre['runner']==d.desc(d.P/'profile.py') and pre['checker']==d.desc(d.P/'profile_preflight.py')
for item in pre['qualification_reports']:q.bind(item['report'])
assert not list(d.P.rglob('__pycache__'))
files=[d.desc(p) for p in sorted(d.P.rglob('*')) if p.is_file()]
d.write(d.P/'admission.json',dict(status='pass',files=files,note='Two immutable matched binaries;64archive qualifications plus2separateHeaptrackpreflights. No formal native/heap captures exist yet.'))
print('Admission PASS;',len(files),'files')
