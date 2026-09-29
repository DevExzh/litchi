"""Reject corrupted retained-report facts without changing captured artifacts."""
import driver as d
import readers as r
import copy,gzip,json
base=d.P/'runs/qualification/00-01-before-all-p0';s=d.read(base/'started.json');original=d.read(base/'report.json')
work=d.SCRATCH/'reader-mutation.json';passed=[]
def check(name,value,case=s['case'],allocation=False):
 work.write_text(json.dumps(value))
 try:r.validate_report(work,case,1,0,allocation=allocation)
 except r.ReplayError:passed.append(name)
 else:raise AssertionError('accepted mutation '+name)
 finally:work.unlink()
mutations=[('schema',lambda x:x.update(schema='wrong')),('sample-count',lambda x:x['samples'].clear()),('member-sha',lambda x:x['corpus']['members'][0].update(sha256='0'*64)),('sequence-sha',lambda x:x['samples'][0]['verification'].update(sequence_sha256='0'*64)),('member-order',lambda x:x['corpus']['members'].reverse()),('cpu-charge',lambda x:x['samples'][0]['resources']['after_operation'].update(cpu_tasks=31)),('retained-worker',lambda x:x['samples'][0]['resources']['after_drop'].update(workers=1)),('disabled-source',lambda x:x['samples'][0]['source_metrics'].update(logical_calls=0)),('negative-time',lambda x:x['samples'][0].update(wall_ns=-1))]
for name,mutate in mutations:
 value=copy.deepcopy(original);mutate(value);check(name,value)
base=d.P/'runs/allocation-preflight/00-01-before-all-p0';s2=d.read(base/'started.json');allocation=d.read(base/'report.json')
for name,mutate in [('unmeasured',lambda x:x['samples'][0]['allocation'].update(status='unavailable')),('live-conservation',lambda x:x['samples'][0]['allocation'].update(live_bytes_after=0)),('region-peak',lambda x:x['samples'][0]['allocation'].update(region_peak_live_bytes=0)),('counter-revision',lambda x:x['metrics'].update(counter_revision='unknown'))]:
 value=copy.deepcopy(allocation);mutate(value);check(name,value,s2['case'],True)
snapshots=json.loads(gzip.decompress((d.P/'runs/memory-preflight/00-01-before-all-p0/snapshots.json.gz').read_bytes()))
value=copy.deepcopy(snapshots);value[0]['pid']+=1
try:r.parse_snapshot_series(value)
except r.ReplayError:passed.append('snapshot-pid')
else:raise AssertionError('accepted snapshot identity mutation')
assert r.nearest_rank([4,1,3,2],.5)==2
assert r.bootstrap_ratio([1]*6,[.5]*6)['estimate']==.5
assert r.bootstrap_ratio([1]*6,[.5]*6)['ci95_low']==.5
assert r.bootstrap_ratio([1]*6,[.5]*6)['ci95_high']==.5
d.write(d.P/'reader-tests.json',dict(status='pass',mutations=passed,statistical_checks=4,reader=d.desc(d.P/'readers.py')))
print('reader tests PASS',len(passed),'mutations')
