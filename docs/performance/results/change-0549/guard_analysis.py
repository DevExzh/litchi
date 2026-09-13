"""Replay public malformed CFB samples with frozen per-row envelopes."""
from pathlib import Path
import hashlib,json,math,statistics
B=Path(__file__).resolve().parent

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def stats(values):
    ordered=sorted(values)
    def q(p):
        position=(len(ordered)-1)*p;lower=math.floor(position);upper=math.ceil(position)
        return ordered[lower]+(ordered[upper]-ordered[lower])*(position-lower)
    return {'p50':q(.5),'p95':q(.95),'p99':q(.99),'mean':statistics.mean(values),'min':min(values),'max':max(values)}

def compute():
    plan=read(B/'plan.json');config=plan['guard'];records={};summaries={};receipts=[]
    for stage in ['baseline','candidate']:
        folder=B/'guard'/stage;identity=read(folder/'binary-guard.json')
        assert sha(folder/'source-manifest.json')==sha(B/stage/'source-manifest.json')
        for repeat in [1,2]:
            execution='candidate' if stage=='baseline' and repeat==2 else stage
            for size in config['sizes']:
                for case in config['cases']:
                    name=f'guard-r{repeat}-{size}-{case}';p=folder/(name+'.json');r=read(folder/(name+'.receipt.json'));d=read(p)
                    assert r['exit_code']==0 and r['binary_sha256']==identity['sha256']
                    assert r['source_manifest_sha256']==identity['source_manifest_sha256']
                    assert r['execution_stage']==execution and r['execution_manifest_sha256']==sha(B/execution/'source-manifest.json')
                    assert r['plan_sha256']==sha(B/'plan.json') and r['script_sha256']==sha(B/'run.py')
                    assert r['command']==['taskset','-c',str(plan['cpu']),identity['path'],'--case',case,'--size',str(size),'--warmup',str(config['warmup']),'--samples',str(config['samples']),'--json',str(p)]
                    for artifact,h in r['artifacts'].items():assert sha(folder/artifact)==h
                    assert d['schema']=='litchi-cfb.perf-chain-guard.v1' and d['case']==case
                    assert d['declared_chain_sectors']==size and d['sector_size']==512 and d['stream_bytes']==size*512
                    assert d['sample_count']==config['samples'] and d['warmup_iterations']==config['warmup']
                    assert d['cleanup_inside_timing'] is False and d['input_clone_outside_timing'] is True
                    assert d['oracle']['expected_error_exact'] and d['oracle']['all_samples_match'] and d['oracle']['valid_reopen_checked']
                    assert d['oracle']['valid_stream_content_checked']==(case=='valid')
                    assert d['expected_error']==d['observed_error']
                    assert (d['expected_error'] is None)==(case=='valid')
                    assert len(d['input_sha256'])==64
                    assert len(d['samples_ns'])==config['samples'] and all(type(v) is int and v>0 for v in d['samples_ns'])
                    key=(stage,repeat,size,case);records[key]=d;summaries[key]=stats(d['samples_ns']);receipts.append({'stage':stage,'repeat':repeat,'size':size,'case':case,'path':str(p.relative_to(B)),'sha256':sha(p),'receipt_sha256':sha(folder/(name+'.receipt.json'))})
    rows=[];adverse=[];drift=[]
    fields=['p50','p95','p99','mean']
    for size in config['sizes']:
        for case in config['cases']:
            baseline_identity={k:v for k,v in records['baseline',1,size,case].items() if k not in ['samples_ns']}
            for stage in ['baseline','candidate']:
                for repeat in [1,2]:assert {k:v for k,v in records[stage,repeat,size,case].items() if k not in ['samples_ns']}==baseline_identity
            for repeat in [1,2]:
                a=summaries['baseline',repeat,size,case];c=summaries['candidate',repeat,size,case];valid=summaries['baseline',repeat,size,'valid']
                change={k:(c[k]/a[k]-1)*100 for k in fields};gates={}
                if case!='valid':
                    for k in ['p50','mean']:
                        gates[k]={'same_invalid_ratio':c[k]/a[k],'baseline_valid_ratio':c[k]/valid[k],'passed':c[k]<=a[k]*config['gates']['invalid_p50_and_mean_max_same_invalid_ratio'] and c[k]<=valid[k]*config['gates']['invalid_p50_and_mean_max_baseline_valid_ratio']}
                rows.append({'size':size,'case':case,'repeat':repeat,'baseline':a,'candidate':c,'change_percent':change,'gates':gates,'passed':all(g['passed'] for g in gates.values())})
                for metric,pct in change.items():
                    if pct>5:adverse.append({'size':size,'case':case,'repeat':repeat,'metric':metric,'change_percent':pct,'baseline':a[metric],'candidate':c[metric]})
            for stage in ['baseline','candidate']:
                a=summaries[stage,1,size,case];c=summaries[stage,2,size,case]
                for metric in fields:
                    pct=(c[metric]/a[metric]-1)*100
                    if abs(pct)>5:drift.append({'stage':stage,'size':size,'case':case,'metric':metric,'change_percent':pct,'repeat1':a[metric],'repeat2':c[metric]})
    return {'status':'pass','admission_passed':all(r['passed'] for r in rows),'plan_sha256':sha(B/'plan.json'),'scope':'Public OleFile::open only; oracle/drop outside clock. Four-stage matched inputs and exacterrors. Allocation unavailable. All threshold/drift flags retained.','statistics':'Sorted linear quantiles at (n-1)*p; arithmetic mean; no timing from instrumented main binary mixed in.','samples':len(records)*config['samples'],'processes':len(records),'receipts':receipts,'rows':rows,'adverse':adverse,'drift':drift}
if __name__=='__main__':print(json.dumps(compute(),indent=2,sort_keys=True))
