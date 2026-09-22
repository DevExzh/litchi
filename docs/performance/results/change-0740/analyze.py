"""Replay the frozen diagnostic callchain classifier over every formal capture."""
import importlib.util,json,statistics
from contract import P,read,sha,guard,projection
spec=importlib.util.spec_from_file_location('stacks',P/'parse-stacks.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
def analyze():
    guard();freeze=read('freeze.json')
    assert all(sha(P/n)==h for n,h in freeze['scripts'].items())
    assert sha(P/'build-fp.json')==freeze['build_sha256']
    receipts=read('formal/manifest.json');assert len(receipts)==6
    rows=[]
    for i,r in enumerate(receipts):
        assert r['exit']==r['script_exit']==0
        case=r['case'];assert r['samples']==(100 if 'plain' in case else 10)
        report=read(f'formal/{i:02d}.json')
        assert projection(report,{'case':case,'samples':r['samples'],'warmups':0,'lane':'native','build':'fp'})==read('oracle.json')[case]
        assert not (P/f'formal/{i:02d}.script.stderr').read_text()
        err=(P/f'formal/{i:02d}.stderr').read_text().lower()
        assert not any(x in err for x in ['failed','error:','corrupt','lost ','throttl'])
        d=mod.analyze((P/f'formal/{i:02d}.stacks').read_text())
        plan=d['buckets']['rooted_plan_prepare'];assert plan['samples']>0
        assert sum(v['period'] for v in d['strict_plan_partition'].values())==plan['period']
        deflate=d['strict_plan_partition'].get('candidate_writer_generated_deflate',{'period':0})['period']
        rows.append({'index':i,'case':case,'generated_deflate_pct_strict_plan_period':100*deflate/plan['period'],**d})
    summary={}
    for case in sorted({r['case'] for r in rows}):
        values=[r['generated_deflate_pct_strict_plan_period'] for r in rows if r['case']==case]
        summary[case]={'processes':len(values),'generated_deflate_pct_strict_plan_period':{'min':min(values),'median':statistics.median(values),'max':max(values)}}
    return {'status':'passed','scope':'diagnostic frame-pointer build; sample period only; unknowns excluded explicitly; no wall-time/causal/removable-cost claim','rows':rows,'summary':summary}
if __name__=='__main__':
    d=analyze();(P/'analysis.json').write_text(json.dumps(d,indent=2,sort_keys=True)+'\n');print(json.dumps(d['summary'],indent=2))
