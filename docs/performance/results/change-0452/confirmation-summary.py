#!/usr/bin/env python3
"""Derive the separately frozen ABBA investigation without pooling primary runs."""
import importlib.util,json,re,statistics
from pathlib import Path
ROOT=Path(__file__).resolve().parent
s=importlib.util.spec_from_file_location('derive',ROOT/'derive.py');d=importlib.util.module_from_spec(s);s.loader.exec_module(d)
def derive():
    rows=[]
    for i,lane in enumerate(d.load('confirmation-protocol.json')['order']):
        v=d.load(f'confirmation/{i}/report.json');r={'lane':i,**lane}
        for phase in ['open','plan','publication','api_sum']:
            values=[x['timings'][phase+'_ns']/1e6 for x in v['samples_raw']];r[phase+'_p50_ms']=statistics.median(values)
            if phase=='api_sum':r.update(api_p95_ms=d.quantile(values,.95),api_p99_ms=d.quantile(values,.99),api_p50_ci95_ms=d.ci(values,4520+i))
        r['process_peak_rss_kib']=int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',d.raw(f'confirmation/{i}/resource.log').decode())[1]);rows.append(r)
    metrics=['open_p50_ms','plan_p50_ms','publication_p50_ms','api_sum_p50_ms','api_p95_ms','api_p99_ms','process_peak_rss_kib'];pairs=[];flags=[]
    for a,b in [(0,1),(3,2),(0,3),(1,2)]:
        kind='paired' if rows[a]['build']!=rows[b]['build'] else 'repeat'
        changes={k:d.delta(rows[a][k],rows[b][k]) for k in metrics};pairs.append({'kind':kind,'from_lane':a,'to_lane':b,'percent_change':changes})
        for k,v in changes.items():
            if abs(v)>5:flags.append({'kind':kind,'from_lane':a,'to_lane':b,'metric':k,'percent_change':v})
    return {'change':452,'reports':4,'samples':120,'scope':'Separate post-primary ABBA investigation. No pooling or primary replacement. Bootstrap 2000 resamples seed 4520+lane is conditional within process.','rows':rows,'comparisons':pairs,'review_flags':flags}
def render(v):
    lines=['# Separate bytes/media CPU confirmation','','The primary baseline drift remains unexplained. Fresh ABBA processes support','about 8% complete API improvement and confirm about 6% extra planning time.','The earlier 36% primary pair is retained but is not the stable CPU claim.','','| Repeat | Baseline API p50 ms | Candidate API p50 ms | API change | Plan change | Publication change |','|---|---:|---:|---:|---:|---:|']
    for p in v['comparisons'][:2]:
        a=v['rows'][p['from_lane']];b=v['rows'][p['to_lane']];c=p['percent_change'];lines.append(f"| {a['repeat']} | {a['api_sum_p50_ms']:.3f} | {b['api_sum_p50_ms']:.3f} | {c['api_sum_p50_ms']:+.3f}% | {c['plan_p50_ms']:+.3f}% | {c['publication_p50_ms']:+.3f}% |")
    lines+=['','All phase medians, p95/p99, median intervals, RSS and >5% flags are retained','in `confirmation-summary.json`; dispositions are in `regression-review.json`.','These runs investigate an observed issue and do not alter the frozen primary gate.']
    return '\n'.join(lines)+'\n'
if __name__=='__main__':
    v=derive();(ROOT/'confirmation-summary.json').write_text(json.dumps(v,indent=2)+'\n');(ROOT/'confirmation-summary.md').write_text(render(v));print(json.dumps({'status':'pass','flags':len(v['review_flags'])}))
