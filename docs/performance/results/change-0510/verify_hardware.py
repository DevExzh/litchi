"""Validate grouped whole-child hardware captures; never operation-local claims."""
import csv, hashlib, json
from pathlib import Path
from tools.summarize_crud_baseline import _validate_elapsed
from tools.validate_perf_corpus_binding import validate_binding

def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def load(path):return json.loads(path.read_text())

def validate(here, builds, reference):
    captures={}
    for stage in ['before','after']:
        for name in ['perf-stat-r1','perf-stat-r2','perf-record']:
            directory=here/stage
            receipt=load(directory/f'{name}-receipt.json')
            assert receipt['exit_code']==0 and receipt['current_source_unchanged']
            assert receipt['binary_sha256']==builds[stage]['binary_sha256']
            assert receipt['source_manifest_sha256']==builds[stage]['source_manifest_sha256']
            assert receipt['command'][:3]==['taskset','-c','2']
            for file,digest in receipt['artifacts'].items():assert sha(directory/file)==digest
            report=load(directory/f'{name}-report.json')
            assert report['binary_identity']['binary_sha256']==builds[stage]['binary_sha256']
            assert report['configuration']['samples_per_case']==1000
            assert report['configuration']['warmup_iterations_per_case']==0
            validate_binding(report,load(directory/f'{name}-catalog.json'))
            assert len(report['results'])==1
            row=report['results'][0]
            _validate_elapsed(row,1000,stage+name)
            for field in ['corpus','sink','output_sha256']:assert row[field]==reference[field]
            if name=='perf-record':
                assert {'perf.data','perf-report.txt','perf-report.log'} <= receipt['artifacts'].keys()
                report_text=(directory/'perf-report.txt').read_text()
                assert '# Total Lost Samples: 0' in report_text
                assert "of event 'cpu/cycles/P'" in report_text
                assert (directory/'perf.data').stat().st_size>0
                continue
            assert name+'.csv' in receipt['artifacts']
            events={}
            with (directory/f'{name}.csv').open() as stream:
                for fields in csv.reader(stream):
                    if not fields or fields[0].startswith('#'):continue
                    assert len(fields)>=5
                    events[fields[2]]={'count':int(fields[0]),'runtime_ns':int(fields[3]),'running_percent':float(fields[4])}
            expected={'cycles','instructions','branches','branch-misses','page-faults','context-switches','cpu-migrations'}
            assert set(events)==expected
            grouped=['cycles','instructions','branches','branch-misses']
            assert all(events[e]['count']>0 and events[e]['running_percent']>=99.9 for e in grouped)
            assert len({events[e]['runtime_ns'] for e in grouped})==1
            assert events['branch-misses']['count']<=events['branches']['count']
            captures[stage,name]=events
    pairs=[]
    for name in ['perf-stat-r1','perf-stat-r2']:
        b,a=[captures[stage,name] for stage in ['before','after']]
        deltas={e:((a[e]['count']/b[e]['count']-1)*100 if b[e]['count'] else None) for e in b}
        pairs.append({'repeat':name,'before':b,'after':a,'change_percent':deltas,'before_whole_child_ipc':b['instructions']['count']/b['cycles']['count'],'after_whole_child_ipc':a['instructions']['count']/a['cycles']['count']})
    return {'profiled_export_samples':6000,'pairs':pairs,'scope':'whole child including fixture creation, opening, validation,1000exports,JSON output; grouped hardware counters include process work outside export clocks; sampling likewise whole-child','cache_counters':'excluded: broader probe had unreliable running time','instrumented_elapsed_excluded_from_native_comparison':True}
