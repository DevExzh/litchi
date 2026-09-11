"""Validate grouped whole-child counters, separately from operation timings."""
import csv,hashlib,json
from tools.summarize_crud_baseline import _validate_elapsed
from tools.validate_perf_corpus_binding import validate_binding

def load(path):return json.loads(path.read_text())
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def validate(here):
    reference=next(r for r in load(here/'before/xls-r1-report.json')['results'] if r['case']=='xls_owned_source_open_one_cell')
    captures={};starts={}
    for stage in ['before','after']:
        directory=here/stage
        build=load(here/'before/build-receipt.json' if stage=='before' else here/'build-receipt.json')
        for repeat in ['r1','r2']:
            name=f'hardware-{repeat}';receipt=load(directory/f'{name}-receipt.json')
            assert receipt['exit_code']==0 and receipt['source_unchanged']
            assert receipt['binary_sha256']==build['binary_sha256']
            assert receipt['source_manifest_sha256']==build['source_manifest_sha256']
            assert receipt['command'][:3]==['taskset','-c','2']
            assert '{cycles,instructions,branches,branch-misses},page-faults,context-switches,cpu-migrations' in receipt['command']
            assert f'{name}.csv' in receipt['artifacts']
            for file,digest in receipt['artifacts'].items():assert sha(directory/file)==digest
            report=load(directory/f'{name}-report.json')
            validate_binding(report,load(directory/f'{name}-catalog.json'))
            assert report['binary_identity']['binary_sha256']==build['binary_sha256']
            assert report['configuration']['samples_per_case']==1000 and report['configuration']['warmup_iterations_per_case']==0
            assert len(report['results'])==1
            row=report['results'][0];_validate_elapsed(row,1000,stage+name)
            for field in ['case','corpus','output_sha256']:assert row[field]==reference[field]
            events={}
            with (directory/f'{name}.csv').open() as stream:
                for fields in csv.reader(stream):
                    if not fields or fields[0].startswith('#'):continue
                    assert len(fields)>=5
                    events[fields[2]]={'count':int(fields[0]),'runtime_ns':int(fields[3]),'running_percent':float(fields[4])}
            assert set(events)=={'cycles','instructions','branches','branch-misses','page-faults','context-switches','cpu-migrations'}
            group=['cycles','instructions','branches','branch-misses']
            assert all(events[e]['count']>0 and events[e]['running_percent']>=99.9 for e in group)
            assert len({events[e]['runtime_ns'] for e in group})==1
            assert events['branch-misses']['count']<=events['branches']['count']
            captures[stage,repeat]=events;starts[stage,repeat]=receipt['started_utc']
    order=[('before','r1'),('after','r1'),('after','r2'),('before','r2')]
    values=[starts[key] for key in order];assert values==sorted(values) and len(set(values))==4
    pairs=[]
    for repeat in ['r1','r2']:
        b,a=[captures[s,repeat] for s in ['before','after']]
        pairs.append({'repeat':repeat,'before':b,'after':a,'change_percent':{e:((a[e]['count']/b[e]['count']-1)*100 if b[e]['count'] else None) for e in b},'before_ipc':b['instructions']['count']/b['cycles']['count'],'after_ipc':a['instructions']['count']/a['cycles']['count']})
    return {'samples':4000,'scope':'whole child including fixture creation, input clones, open/query, oracles, drops and JSON reporting; not operation-local IPC or counters','instrumented_timings_excluded':True,'cache_events':'not collected; prior broad probe was unreliable','pairs':pairs}
