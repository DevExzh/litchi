"""Replay the frozen matrix and summarize processes without pooling samples."""
import json, math, random, statistics
from run import P, PHASES, read, sha, guard, schedule, command, projection, verify_freeze

def stats(xs):
    ys=sorted(xs); n=len(ys)
    return {'p50':statistics.median(ys),'mean':statistics.mean(ys),'p95':ys[math.ceil(.95*n)-1],'p99':ys[math.ceil(.99*n)-1],'maximum':ys[-1]}

def summary(xs):
    rng=random.Random(7339)
    boot=sorted(statistics.median(rng.choices(xs,k=len(xs))) for _ in range(10000))
    return {'median':statistics.median(xs),'minimum':min(xs),'maximum':max(xs),'bootstrap_median_95':[boot[249],boot[9749]],'values':xs}

def analyze():
    guard();verify_freeze(); rows=schedule(); saved=read('captures/manifest.json');assert len(saved)==len(rows)
    processes=[];previous=0
    expected_files={'manifest.json'}
    for i,(row,record) in enumerate(zip(rows,saved)):
        assert all(record[k]==v for k,v in row.items())
        assert record['command']==command(row,P/'captures'/f'{i:03d}.json')
        assert record['exit']==0 and previous<=record['monotonic_start']<record['monotonic_end']
        previous=record['monotonic_end']; assert record['started']<=record['ended']
        names=[f'captures/{i:03d}.{ext}' for ext in ['json','stdout','stderr']]
        assert set(record['files'])==set(names)
        for n,h in record['files'].items(): assert sha(P/n)==h,n
        expected_files.update(f'{i:03d}.{ext}' for ext in ['json','stdout','stderr'])
        assert (P/f'captures/{i:03d}.stderr').read_bytes()==b''
        report=read(f'captures/{i:03d}.json');assert projection(report,row)==read('oracle.json')[row['case']]
        r=report['results'][0]; x=r['source']['pptx_cross_copy']
        item=row.copy()
        if row['lane']=='native':
            ordered=[0]*row['samples']
            for index,value in zip(r['elapsed_ns']['sample_order'],x['lifecycle_ns']): ordered[index]=value
            values={p:x[p] for p in PHASES}
            values['unassigned_ns']=[total-sum(x[p][i] for p in ['plan_ns','commit_ns','publication_ns']) for i,total in enumerate(x['lifecycle_ns'])]
            item['stats']={p:stats(v) for p,v in values.items()}
            item['phase_share_percent']={p:statistics.median([100*v/t for v,t in zip(values[p],x['lifecycle_ns'])]) for p in ['plan_ns','commit_ns','publication_ns','unassigned_ns']}
            item['ordered_lifecycle_ns']=ordered
            item['last10_vs_first10_percent']=(statistics.median(ordered[-10:])/statistics.median(ordered[:10])-1)*100
        else:
            item['allocations']=r['operation_metrics']['allocation']
        processes.append(item)
    assert {p.name for p in (P/'captures').iterdir()}==expected_files
    groups=[]
    for case in read('plan.json')['cases']:
        ps=[p for p in processes if p['lane']=='native' and p['case']==case]
        metrics={phase:{metric:summary([p['stats'][phase][metric] for p in ps]) for metric in ['p50','mean','p95','p99','maximum']} for phase in ps[0]['stats']}
        flags=[]
        for phase,m in metrics.items():
            for metric,s in m.items():
                spread=(s['maximum']/s['minimum']-1)*100
                if spread>5:flags.append({'phase':phase,'metric':metric,'across_process_spread_percent':spread})
        groups.append({'case':case,'metrics_ns':metrics,'phase_share_percent':{phase:summary([p['phase_share_percent'][phase] for p in ps]) for phase in ps[0]['phase_share_percent']},'spread_flags':flags,'order_flags':[{'repeat':p['repeat'],'change_percent':p['last10_vs_first10_percent']} for p in ps if abs(p['last10_vs_first10_percent'])>5]})
    return {'status':'passed','native_processes':18,'native_samples':540,'allocation_processes':6,'processes':processes,'groups':groups,'scope':'Current source descriptive process baseline. Phase wall-clock fractions are not CPU fractions or causal removable costs. Reopen is outside lifecycle; unassigned includes ingress/snapshots and clock/loop overhead. No speedup or historical comparison.'}

if __name__=='__main__':
    result=analyze();(P/'analysis.json').write_text(json.dumps(result,indent=2)+'\n');print('PASS24 processes,540 native samples,exact qualification oracles and all diagnostics')
