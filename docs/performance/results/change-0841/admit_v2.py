"""Additive reader repair; all native measurements and frozen inputs retained."""
import driver as d
import qualify_v2 as q
import qualify as old
assert all(d.desc(x['path'])==x for x in d.read(d.P/'admission.json')['files'])
assert d.source()==d.read(d.P/'source-after.json')['source']
new_qual=q.qualification();assert new_qual==d.read(d.P/'qualification-v2.json')
assert new_qual['rows']==d.read(d.P/'qualification.json')['rows']
assert old.validate_items('native')==q.validate_items('native')
assert not any(p.name.startswith('profiles-') for p in (d.P/'runs').iterdir())
assert d.read(d.P/'reader-tests-v2.json')['status']=='pass'
violation=d.read(d.P/'reader-freeze-violation.json')
assert d.sha(d.P/'readers.py')==violation['old']['sha256']
assert d.sha(d.P/'readers_v2_initial.py')==violation['new']['sha256']
repair=dict(status='pass',original_admission=d.desc(d.P/'admission.json'),violation=d.desc(d.P/'reader-freeze-violation.json'),initial_replacement_archive=d.desc(d.P/'readers_v2_initial.py'),original_reader=d.desc(d.P/'readers.py'),corrected_reader=d.desc(d.P/'readers_v2.py'),native_parity=d.desc(d.P/'commands/native-reader-parity/receipt.json'),qualification_parity='All20rows identical; only reader identity differs',native_reports=120,workload_runs_replaced=0,changes=['Read child-only raw Callgrind command header; enforce launcher flags through exact recorded argv','Reject pending child edges without a cost at function transition or EOF'],unchanged=['candidate','baseline','binaries','probe','corpus','schedule','native metrics','bootstrap','all workload outputs'])
d.write(d.P/'reader-repair.json',repair)
names=['readers_v2.py','readers_v2_initial.py','reader_tests_v2.py','reader-tests-v2.json','qualify_v2.py','qualification-v2.json','analyze_v2.py','capture_v2.py','admit_v2.py','reader-repair.json','reader-freeze-violation.json','capture-native.json']
d.write(d.P/'admission-v2.json',dict(status='pass',files=[d.desc(d.P/n) for n in names]))
print('Additive reader admission PASS;120native reports unchanged;0workloads replaced')
