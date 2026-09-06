#!/usr/bin/env python3
"""Derive all matched simulation observations; never pool configurations or repeats."""
import argparse,gzip,json,math,random,statistics
from pathlib import Path
ROOT=Path(__file__).resolve().parent

def load(path):return json.loads(path.read_text())
def raw(path):return path.read_bytes() if path.exists() else gzip.decompress(Path(str(path)+'.gz').read_bytes())
def stats(values,seed=448):
    v=sorted(values);p=lambda q:v[min(len(v)-1,math.ceil(q*len(v))-1)]
    rng=random.Random(seed);boot=sorted(statistics.median(rng.choices(v,k=len(v))) for _ in range(2000))
    return {'samples':values,'p50':statistics.median(v),'p95':p(.95),'p99':p(.99),'mean':statistics.mean(v),'min':v[0],'max':v[-1],'p50_bootstrap_95':[boot[49],boot[1949]],'bootstrap_scope':'2000 deterministic within-process resamples; not machine/day uncertainty'}
def relative(old,new):return 100*(new/old-1) if old else (0 if new==0 else None)
def derive():
    protocol=load(ROOT/'protocol.json');rows=[]
    for i,lane in enumerate(protocol['order']):
        report=load(ROOT/f'runs/{i}/report.json');samples=report['samples_raw']
        counters={}
        for phase in ['opened','planned','published']:
            counters[phase]={}
            for field in ['logical_calls','requested_bytes','returned_bytes','transfer_paced_calls','transfer_delay_ns']:
                values=[]
                for sample in samples:
                    point=next(p for p in sample['phases'] if p['label']==phase)
                    values.append(sum(point[owner]['delta'][field] for owner in ['source_reads','destination_reads']))
                counters[phase][field]=stats(values)
        peak=[]
        for sample in samples:peak.append(max(p['source_budget']['memory_used']+p['destination_budget']['memory_used'] for p in sample['phases']))
        import re
        rss=int(re.search(rb'Maximum resident set size \(kbytes\):\s*(\d+)',raw(ROOT/f'runs/{i}/resource.log'))[1])*1024
        rows.append({'lane':i,**lane,'timings':{k:stats([v['timings'][k] for v in samples]) for k in ['open_source_ns','open_destination_ns','open_ns','plan_ns','publication_ns','api_sum_ns']},'phase_read_deltas':counters,'managed_boundary_memory':stats(peak),'whole_process_peak_rss_bytes':rss,'output_sha256':report['expected_output_sha256'],'output_bytes':report['expected_output_bytes'],'source_sha256':report['source_archive_sha256'],'destination_sha256':report['destination_archive_sha256']})
    repeats=[];pairs=[]
    for corpus in ['plain','media-rich']:
        for minimum_service in [False,True]:
            r1=next(r for r in rows if (r['corpus'],r['minimum_service'],r['repeat'])==(corpus,minimum_service,'R1'));r2=next(r for r in rows if (r['corpus'],r['minimum_service'],r['repeat'])==(corpus,minimum_service,'R2'))
            for phase in ['open_ns','plan_ns','publication_ns','api_sum_ns']:
                for metric in ['p50','p95','p99']:
                    change=relative(r1['timings'][phase][metric],r2['timings'][phase][metric]);repeats.append({'corpus':corpus,'minimum_service':minimum_service,'metric':phase+'.'+metric,'relative_percent':change,'flagged':change is None or abs(change)>5})
            change=relative(r1['whole_process_peak_rss_bytes'],r2['whole_process_peak_rss_bytes']);repeats.append({'corpus':corpus,'minimum_service':minimum_service,'metric':'whole_process_peak_rss_bytes','relative_percent':change,'flagged':abs(change)>5})
        for repeat in ['R1','R2']:
            a=next(r for r in rows if (r['corpus'],r['minimum_service'],r['repeat'])==(corpus,False,repeat));b=next(r for r in rows if (r['corpus'],r['minimum_service'],r['repeat'])==(corpus,True,repeat))
            for phase in ['open_ns','plan_ns','publication_ns','api_sum_ns']:
                for metric in ['p50','p95','p99']:
                    pairs.append({'corpus':corpus,'repeat':repeat,'metric':phase+'.'+metric,'separate':a['timings'][phase][metric],'minimum_service':b['timings'][phase][metric],'relative_percent':relative(a['timings'][phase][metric],b['timings'][phase][metric]),'scope':'minimum-service versus separate-sleep model calibration; not production regression'})
            pairs.append({'corpus':corpus,'repeat':repeat,'metric':'whole_process_peak_rss_bytes','separate':a['whole_process_peak_rss_bytes'],'minimum_service':b['whole_process_peak_rss_bytes'],'relative_percent':relative(a['whole_process_peak_rss_bytes'],b['whole_process_peak_rss_bytes']),'scope':'whole-process RSS, not operation allocation attribution'})
    gate=[p for p in pairs if p['corpus']=='plain' and p['metric']=='api_sum_ns.p50'];assert len(gate)==2
    return {'acceptance':{'plain_p50_gate':all(p['relative_percent']<=-5 for p in gate),'scope':'frozen 5% timer-model gate; correctness and all paired/repeat reviews additionally required'},'change':448,'rows':rows,'repeat_review':repeats,'comparisons':pairs,'operation_allocator':'unavailable; neither managed boundary gauges nor process RSS substitute for operation allocation attribution','scope':'matched single-build delay-policy calibration, 8 reports/240 samples; no actual network/cold/scaling/production speedup claim'}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();v=derive();path=ROOT/'summary.json'
    if a.check:assert load(path)==v
    else:
        with path.open('x') as f:f.write(json.dumps(v,indent=2)+'\n')
    print('VALID')
