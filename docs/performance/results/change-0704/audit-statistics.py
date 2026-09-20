#!/usr/bin/env python3
"""Independently recompute initial native and allocation headline statistics."""
import json, math, statistics
from pathlib import Path
P=Path(__file__).resolve().parent

def read(name): return json.loads((P/name).read_text())
def rows(name):
    lines=(P/name).read_text().splitlines()
    header=next(x.split('\t') for x in lines if x.startswith('sample\t'))
    return [dict(zip(header,map(int,x.split('\t')))) for x in lines if x.split('\t')[0].isdigit()]

raw={}
for filename in ['native-runs-baseline.json','native-runs-compare.json']:
    for r in read(filename):raw[r['case'],r['leg']]=rows(r['output'])
summary={}
for r in read('native-summary.json'):
    v=sorted(x[r['phase']] for x in raw[r['case'],r['leg']])
    expected={'samples':len(v),'p50_ns':statistics.median(v),'mean_ns':statistics.mean(v),
              'p95_ns':v[math.ceil(len(v)*.95)-1],'p99_ns':v[math.ceil(len(v)*.99)-1],
              'min_ns':v[0],'max_ns':v[-1]}
    assert all(r[k]==value for k,value in expected.items()),r
    key=(r['case'],r['leg'],r['phase']);assert key not in summary
    summary[key]=r
assert len(summary)==13*6*6
for r in read('native-comparisons.json'):
    b=summary[r['case'],r['baseline_leg'],r['phase']]
    c=summary[r['case'],r['candidate_leg'],r['phase']]
    for k in ['p50_ns','mean_ns','p95_ns','p99_ns']:
        assert r['candidate'][k]==c[k] and r['baseline'][k]==b[k]
        assert math.isclose(r['delta_ns'][k],c[k]-b[k],abs_tol=1e-8)
        assert math.isclose(r['delta_pct'][k],100*(c[k]-b[k])/b[k],abs_tol=1e-8)
alloc={}
for phase in ['baseline','candidate']:
    for r in read('allocation-runs-'+phase+'.json'):alloc[r['case'],phase]=rows(r['output'])
for r in read('allocation-summary.json'):
    source=alloc[r['case'],r['run_phase']]; phase=r['phase']
    for field,values in r['metrics'].items():assert values==[x[phase+'_'+field] for x in source]
    assert r['peak_above_start']==[x[phase+'_peak_live_bytes']-x[phase+'_baseline_live_bytes'] for x in source]
    assert r['net_live_change']==[x[phase+'_current_live_bytes']-x[phase+'_baseline_live_bytes'] for x in source]
for r in read('allocation-comparisons.json'):
    for field,delta in r['delta'].items():
        assert delta==[c-b for c,b in zip(r['candidate']['metrics'][field],r['baseline']['metrics'][field])]
print('PASS: independently recomputed 468 native summaries, 312 comparisons and 156 allocation summaries')
refusal={}
bindings=read('refusal-bindings.json')
for file in ['refusal-runs-baseline.json','refusal-runs-compare.json']:
    for run in read(file):
        current=None; values={}; metadata={}
        for line in (P/run['output']).read_text().splitlines():
            parts=line.split('\t')
            if parts[0]=='case':current=parts[1];values[current]=[];metadata[current]={}
            elif parts[0].isdigit():
                assert int(parts[0])==len(values[current]);values[current].append(int(parts[1]))
            elif current and parts[0]!='sample_ns':metadata[current][parts[0]]=parts[1:]
        assert len(values)==10
        for case,v in values.items():
            assert len(v)==100
            for key,value in bindings[case].items():assert metadata[case][key]==value
            refusal[case,run['leg']]=sorted(v)
refusal_stats={}
for r in read('refusal-summary.json'):
    v=refusal[r['case'],r['leg']]
    expected={'samples':len(v),'p50_ns':statistics.median(v),'mean_ns':statistics.mean(v),
              'p95_ns':v[94],'p99_ns':v[98]}
    assert all(r[k]==value for k,value in expected.items())
    refusal_stats[r['case'],r['leg']]=r
assert len(refusal_stats)==60
for r in read('refusal-comparisons.json'):
    cleg,bleg=r['pair'].split('/'); c=refusal_stats[r['case'],cleg];b=refusal_stats[r['case'],bleg]
    for k in ['p50_ns','mean_ns','p95_ns','p99_ns']:
        assert math.isclose(r['delta_ns'][k],c[k]-b[k],abs_tol=1e-8)
        assert math.isclose(r['delta_pct'][k],100*(c[k]-b[k])/b[k],abs_tol=1e-8)
print('PASS: independently recomputed 60 refusal summaries, exact fixture/error bindings and pair comparisons')
