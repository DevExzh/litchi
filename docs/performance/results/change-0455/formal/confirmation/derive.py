#!/usr/bin/env python3
"""Derive separate timing/allocation populations and exact logical work deltas."""
import hashlib,json,random,re,statistics
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def load(p):return json.loads(p.read_text())
def q(v,p):
    v=sorted(v);i=(len(v)-1)*p;a=int(i);return v[a]+(v[min(a+1,len(v)-1)]-v[a])*(i-a)
def stats(v):return {'p50':q(v,.5),'p95':q(v,.95),'p99':q(v,.99)}
def change(a,b):return (b/a-1)*100 if a else None
def ci(a,b,seed):
    rng=random.Random(seed);v=[change(statistics.median(rng.choices(a,k=len(a))),statistics.median(rng.choices(b,k=len(b)))) for _ in range(2000)];return [q(v,.025),q(v,.975)]
def work(row):
    return {'phases':{p['label']:{k:p[k] for k in ['source_reads','destination_reads','source_budget','destination_budget']} for p in row['phases']},'publication_sink':row['publication_sink']}
def derive():
    protocol=load(ROOT/'protocol.json');lanes=[]
    for index,lane in enumerate(protocol['lanes']):
        r=ROOT/'runs'/str(index);report=load(r/'report.json');rows=report['samples_raw'];assert len(rows)==protocol['samples']
        logical=work(rows[0]);assert all(work(row)==logical for row in rows),'nonconstant logical work'
        values={k:[row['timings'][k] for row in rows] for k in rows[0]['timings'] if k.endswith('_ns')}
        allocation={}
        if lane['instrumentation']=='allocator':
            for phase in ['open_source','open_destination','plan','publication']:
                key=phase+'_allocation_metrics';allocation[phase]={k:stats([row['timings'][key][k] for row in rows]) for k,v in rows[0]['timings'][key].items() if type(v) is int}
        rss=int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',(r/'resource.log').read_text())[1])
        lanes.append({'index':index,**lane,'timings_ns':{k:stats(v) for k,v in values.items()},'values_ns':values,'allocation':allocation,'rss_kib':rss,'work':logical,'output_sha256':report['expected_output_sha256'],'output_bytes':report['expected_output_bytes']})
    pairs=[];flags=[]
    for b in lanes:
        if b['build']!='candidate':continue
        a=next(x for x in lanes if x['build']=='baseline' and all(x[k]==b[k] for k in ['provider','corpus','repeat','instrumentation']))
        assert (a['output_sha256'],a['output_bytes'])==(b['output_sha256'],b['output_bytes'])
        pair={k:b[k] for k in ['provider','corpus','repeat','instrumentation']};pair.update(baseline_lane=a['index'],candidate_lane=b['index'])
        pair['timing_percent']={k:{p:change(a['timings_ns'][k][p],b['timings_ns'][k][p]) for p in ['p50','p95','p99']} for k in b['timings_ns']}
        pair['api_median_change_bootstrap_95_percent']=ci(a['values_ns']['api_sum_ns'],b['values_ns']['api_sum_ns'],455+b['index'])
        pair['rss_percent']=change(a['rss_kib'],b['rss_kib'])
        pair['allocation_percent']={phase:{k:{p:change(a['allocation'][phase][k][p],v) for p,v in vals.items()} for k,vals in fields.items()} for phase,fields in b['allocation'].items()}
        pair['publication_destination_read_call_delta']=b['work']['phases']['published']['destination_reads']['delta']['logical_calls']-a['work']['phases']['published']['destination_reads']['delta']['logical_calls']
        pair['publication_sink_write_call_delta']=b['work']['publication_sink']['write_calls']-a['work']['publication_sink']['write_calls']
        for phase in a['work']['phases']:
            for provider in ['source_reads','destination_reads']:
                av=a['work']['phases'][phase][provider];bv=b['work']['phases'][phase][provider]
                assert av['returned_bytes']==bv['returned_bytes'],'transferred bytes changed'
                assert av['requested_bytes']==bv['requested_bytes'],'requested bytes changed'
                if phase in ['baseline','opened','planned'] or provider=='source_reads':assert av==bv,'unrelated read work changed'
            for owner in ['source_budget','destination_budget']:assert a['work']['phases'][phase][owner]==b['work']['phases'][phase][owner],'managed resource counters changed'
        if b['instrumentation']=='normal':
            for metric,values in pair['timing_percent'].items():
                for percentile,value in values.items():
                    if abs(value)>5:flags.append({'pair':{k:pair[k] for k in ['provider','corpus','repeat']},'metric':metric+'.'+percentile,'percent':value})
            if abs(pair['rss_percent'])>5:flags.append({'pair':{k:pair[k] for k in ['provider','corpus','repeat']},'metric':'rss_kib','percent':pair['rss_percent']})
        pairs.append(pair)
    for lane in lanes:del lane['values_ns']
    return {'change':455,'samples':sum(protocol['samples'] for _ in lanes),'method':'linear-interpolated quantiles; independent within-lane median bootstrap, 2000 resamples, seed 455+candidate lane; descriptive two-repeat ABBA, not a population-tail guarantee','lanes':lanes,'comparisons':pairs,'review_flags':flags}
def main():
    result=derive();(ROOT/'measurements.json').write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
    lines=['# 0455 matched transfer-chunk measurements','',result['method'],'','| Provider | Corpus | Repeat | API p50 change | Bootstrap 95% interval | Publication p50 change | RSS change |','|---|---|---|---:|---|---:|---:|']
    for p in result['comparisons']:
        if p['instrumentation']!='normal':continue
        interval=p['api_median_change_bootstrap_95_percent'];lines.append(f"| {p['provider']} | {p['corpus']} | {p['repeat']} | {p['timing_percent']['api_sum_ns']['p50']:+.3f}% | [{interval[0]:+.3f}%, {interval[1]:+.3f}%] | {p['timing_percent']['publication_ns']['p50']:+.3f}% | {p['rss_percent']:+.3f}% |")
    lines+=['','All absolute >5% flags (including improvements) follow. Positive values are regressions.','']
    for flag in result['review_flags']:lines.append('- '+json.dumps(flag,sort_keys=True))
    lines+=['','Complete phase percentiles, separate allocator populations, request histograms, sink counters and resource counters are retained in `measurements.json`.']
    (ROOT/'measurements.md').write_text('\n'.join(lines)+'\n');print(json.dumps({'status':'pass','samples':result['samples'],'review_flags':len(result['review_flags'])}))
if __name__=='__main__':main()
