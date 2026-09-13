"""Recompute all summaries without modifying captured samples."""
from pathlib import Path
import json, math, statistics
B=Path(__file__).resolve().parent
METRICS=['p50','mean','p95','p99']
def metrics(values):
    s=sorted(values)
    return {'p50':statistics.median(s),'mean':statistics.mean(s),'p95':s[math.ceil(.95*len(s))-1],'p99':s[math.ceil(.99*len(s))-1]}
def compute():
    freeze=json.loads((B/'freeze.json').read_text());raw={};rows=[];reviews=[];drift=[]
    for f in freeze['fixtures']:
        for v in ['baseline','counted']:
            for r in [1,2]:
                d=json.loads((B/f'native-{v}-{r}-{f}.json').read_text());assert d['samples']==len(d['duration_ns'])==100
                raw[f,v,r]=metrics(d['duration_ns'])
        for r in [1,2]:
            a,c=raw[f,'baseline',r],raw[f,'counted',r]
            delta={m:100*(c[m]/a[m]-1) for m in METRICS}
            rows.append({'fixture':f,'repeat':r,'baseline_ns':a,'counted_ns':c,'change_percent':delta})
            for m in METRICS:
                if delta[m]>5:
                    reason={'sparse':'Sparse-control adverse result must remain explicit before any workflow integration.','early-reject':'Early-exit adverse result must remain explicit before any workflow integration.'}.get(f,'Adverse tail retained; cannot infer workflow benefit from scanner median.')
                    reviews.append({'kind':'adverse','fixture':f,'repeat':r,'metric':m,'percent':delta[m],'review':reason})
        for v in ['baseline','counted']:
            delta={m:100*(raw[f,v,2][m]/raw[f,v,1][m]-1) for m in METRICS}
            drift.append({'fixture':f,'variant':v,'repeat2_vs_repeat1_percent':delta})
            for m in METRICS:
                if abs(delta[m])>5:reviews.append({'kind':'drift','fixture':f,'variant':v,'metric':m,'percent':delta[m],'review':'Repeat variability retained without rerun or sample filtering. Limits precise magnitude and tail claims; does not establish workflow admission.'})
    return {'scope':'Isolated hot scanner only. All raw samples retained; nearest-rank tails, median and arithmetic mean. Two process repeats describe variability, not a formal population confidence interval. No workflow or memory speedup claim.','rows':rows,'drift':drift,'reviews':reviews,'decision':'Isolated screening only; integrated admission requires fresh original gates.'}
if __name__=='__main__':(B/'summary.json').write_text(json.dumps(compute(),indent=2)+'\n')
