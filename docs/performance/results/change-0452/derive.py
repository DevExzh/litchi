#!/usr/bin/env python3
"""Derive complete lifecycle comparisons and all >5% latency/RSS review flags."""
import gzip,json,random,re,statistics
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def require(ok,message):
    if not ok:raise ValueError(message)
def raw(name):
    p=ROOT/name
    return p.read_bytes() if p.exists() else gzip.decompress(p.with_name(p.name+'.gz').read_bytes())
def load(name):return json.loads(raw(name))
def quantile(v,q):
    v=sorted(v);x=(len(v)-1)*q;i=int(x);return v[i]+(v[min(i+1,len(v)-1)]-v[i])*(x-i)
def ci(values,seed):
    rng=random.Random(seed);b=[statistics.median(rng.choices(values,k=len(values))) for _ in range(2000)]
    return [quantile(b,.025),quantile(b,.975)]
def delta(a,b):return (b/a-1)*100
def derive():
    protocol=load('protocol.json');rows=[]
    for i,lane in enumerate(protocol['order']):
        report=load(f'runs/{i}/report.json');samples=report['samples_raw'];require(len(samples)==30,'30 samples')
        row={'lane':i,**lane}
        for phase in ['open','plan','publication','api_sum']:
            v=[s['timings'][phase+'_ns']/1e6 for s in samples]
            row[phase+'_p50_ms']=statistics.median(v)
            if phase=='api_sum':
                row['api_p95_ms']=quantile(v,.95);row['api_p99_ms']=quantile(v,.99);row['api_p50_ci95_ms']=ci(v,452+i)
        match=re.search(r'Maximum resident set size \(kbytes\): (\d+)',raw(f'runs/{i}/resource.log').decode());require(match is not None,'process RSS')
        row['process_peak_rss_kib']=int(match[1])
        phases={p['label']:p for p in samples[0]['phases']}
        row['source_work_units']=phases['published']['source_cache']['budget_work_used']-phases['opened']['source_cache']['budget_work_used']
        row['source_planned_reserved_bytes']=phases['planned']['source_cache']['budget_memory_used']
        row['source_published_reserved_bytes']=phases['published']['source_cache']['budget_memory_used']
        row['source_after_plan_drop_reserved_bytes']=phases['drop_plan']['source_cache']['budget_memory_used']
        row['source_publication_cache_hits']=phases['published']['source_cache']['hits']-phases['planned']['source_cache']['hits']
        row['source_publication_cold_loads']=phases['published']['source_cache']['cold_loads']-phases['planned']['source_cache']['cold_loads']
        row['io']={}
        for label in ['opened','planned','published']:
            for owner in ['source','destination']:
                point=phases[label][owner+'_reads'];v=point['delta']
                row['io'][label+'_'+owner]=None if v is None else {k:v[k] for k in ['logical_calls','returned_bytes']}
                for sample in samples:
                    p=next(p for p in sample['phases'] if p['label']==label)[owner+'_reads']['delta']
                    require((None if p is None else {k:p[k] for k in ['logical_calls','returned_bytes']})==row['io'][label+'_'+owner],'deterministic I/O per sample')
        rows.append(row)
    pairs=[];flags=[];repeats=[]
    metrics=['open_p50_ms','plan_p50_ms','publication_p50_ms','api_sum_p50_ms','api_p95_ms','api_p99_ms','process_peak_rss_kib']
    for provider in ['bytes','range']:
        for corpus in ['plain','media-rich']:
            for repeat in ['R1','R2']:
                a=next(r for r in rows if (r['provider'],r['corpus'],r['repeat'],r['build'])==(provider,corpus,repeat,'baseline'))
                b=next(r for r in rows if (r['provider'],r['corpus'],r['repeat'],r['build'])==(provider,corpus,repeat,'candidate'))
                values={k:delta(a[k],b[k]) for k in metrics}
                pair={'provider':provider,'corpus':corpus,'repeat':repeat,'baseline_lane':a['lane'],'candidate_lane':b['lane'],'percent_change':values}
                pairs.append(pair)
                for k,v in values.items():
                    if abs(v)>5:flags.append({'kind':'paired','provider':provider,'corpus':corpus,'repeat':repeat,'metric':k,'percent_change':v})
                for label in ['opened','planned','published']:require(a['io'][label+'_destination']==b['io'][label+'_destination'],'destination I/O equality')
                if provider=='range' and corpus=='media-rich':
                    require(values['api_sum_p50_ms']<=-protocol['acceptance']['range_media_p50_reduction_percent'],'frozen range median gate')
                    old=sum(a['io'][label+'_source']['returned_bytes'] for label in ['opened','planned','published'])
                    new=sum(b['io'][label+'_source']['returned_bytes'] for label in ['opened','planned','published'])
                    require(old-new>=16*1024*1024,'one large media source pass removed')
    for provider in ['bytes','range']:
        for corpus in ['plain','media-rich']:
            for build in ['baseline','candidate']:
                a=next(r for r in rows if (r['provider'],r['corpus'],r['build'],r['repeat'])==(provider,corpus,build,'R1'))
                b=next(r for r in rows if (r['provider'],r['corpus'],r['build'],r['repeat'])==(provider,corpus,build,'R2'))
                values={k:delta(a[k],b[k]) for k in metrics}
                repeats.append({'provider':provider,'corpus':corpus,'build':build,'percent_change':values})
                for k,v in values.items():
                    if abs(v)>5:flags.append({'kind':'repeat','provider':provider,'corpus':corpus,'build':build,'metric':k,'percent_change':v})
    return {'change':452,'formal_reports':16,'formal_samples':480,'quantiles':'linear interpolation on sorted samples; p50 median','confidence_interval':'2000 independent bootstrap resamples of each lane median, seed 452+lane; process samples share fixture generation and allocator state','rows':rows,'pairs':pairs,'repeats':repeats,'review_flags':flags}
def render(v):
    lines=['# Matched PPTX capture lifecycle measurements','','Each row is a fresh CPU-2 process: 3 warmups and 30 samples. API time includes','open, planning and publication. RSS is whole-process high water, including','untimed corpus/gates/reporting; it is not a timed allocator peak.','','| Provider | Corpus | Repeat | Baseline p50 ms | Candidate p50 ms | Change | Baseline RSS KiB | Candidate RSS KiB |','|---|---|---|---:|---:|---:|---:|---:|']
    for p in v['pairs']:
        a=v['rows'][p['baseline_lane']];b=v['rows'][p['candidate_lane']]
        lines.append(f"| {p['provider']} | {p['corpus']} | {p['repeat']} | {a['api_sum_p50_ms']:.3f} | {b['api_sum_p50_ms']:.3f} | {p['percent_change']['api_sum_p50_ms']:+.3f}% | {a['process_peak_rss_kib']} | {b['process_peak_rss_kib']} |")
    lines+=['','Raw p95/p99, median bootstrap intervals, per-phase medians, all read-work','counters, reserved-memory gauges and every >5% paired/repeat flag are in','`measurements.json`. See `regression-review.md` for interpretation.','','Range simulation uses 65,536-byte maximum returns, 200 microseconds per read,','and 25 MiB/s separate-sleep transfer pacing. It is not an actual network or','cold-filesystem measurement. Registry/default and native coverage are unchanged.']
    return '\n'.join(lines)+'\n'
if __name__=='__main__':
    v=derive();(ROOT/'measurements.json').write_text(json.dumps(v,indent=2)+'\n');(ROOT/'measurements.md').write_text(render(v));print(json.dumps({'status':'pass','reports':16,'samples':480,'review_flags':len(v['review_flags'])}))
