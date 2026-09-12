"""Validate all three allocation regions and the frozen planning allocation gate."""
import json
import statistics
from pathlib import Path
import analyze as native

HERE = Path(__file__).resolve().parent
FIELDS = ('allocation_calls','reallocation_calls','allocated_bytes','incremental_region_peak_live_bytes')

def analyze():
    plan = json.loads((HERE/'plan.json').read_text())
    gates = json.loads((HERE/'allocation-gates.json').read_text())
    assert gates['planning_allocation_calls_reduction_percent'] == plan['gates']['planning_allocation_calls_reduction_percent']
    jobs = native._expected_allocation_jobs(plan)
    evidence = {}
    for stage in ['baseline','candidate']:
        checked = native._check_stage(stage, plan, jobs, True)
        evidence[stage] = {}
        for job in jobs:
            report = json.loads((HERE/stage/(job['name']+'.json')).read_text())
            result = report['results'][0]
            source = result['source']['xlsx_cell_values']
            phases = {}
            for phase in ['plan','commit','publication']:
                samples = source[phase+'_allocation_metrics']
                assert len(samples) == len(source[phase+'_ns']) == job['samples']
                parsed = [native.BASE.allocation_sample(s,stage+'/'+job['name']+'/'+phase,True) for s in samples]
                phases[phase] = {field:[s[field] for s in parsed] for field in FIELDS}
            evidence[stage][job['name']] = dict(phases=phases,
                identity=next(r['identity'] for r in checked['allocation']['rows'] if r['name']==job['name']))
    rows = []
    adverse = []
    decisions = []
    for job in jobs:
        name = job['name']
        a,b = evidence['baseline'][name],evidence['candidate'][name]
        assert a['identity'] == b['identity'],name
        for phase in ['plan','commit','publication']:
            metrics = {}
            for field in FIELDS:
                av,bv = a['phases'][phase][field],b['phases'][phase][field]
                am,bm = statistics.median(av),statistics.median(bv)
                change = (bm/am-1)*100 if am else (0 if bm==0 else None)
                metrics[field] = dict(baseline=av,candidate=bv,baseline_p50=am,candidate_p50=bm,change_percent=change)
                if change is None or change > 5:
                    adverse.append(dict(name=name,phase=phase,field=field,change_percent=change))
            rows.append(dict(name=name,shape=job['shape'],repeat=job['repeat'],phase=phase,metrics=metrics))
        # Demand a reduction even for the least favorable sampled pair.
        before = min(a['phases']['plan']['allocation_calls'])
        after = max(b['phases']['plan']['allocation_calls'])
        assert before > 0
        reduction = 100*(before-after)/before
        bytes_nonincreasing = max(b['phases']['plan']['allocated_bytes']) <= min(a['phases']['plan']['allocated_bytes'])
        decisions.append(dict(name=name,worst_pair_calls_reduction_percent=reduction,
            bytes_nonincreasing=bytes_nonincreasing,
            passed=reduction >= gates['planning_allocation_calls_reduction_percent'] and bytes_nonincreasing))
    return dict(status='pass',planning_gate_passed=all(d['passed'] for d in decisions),
                decisions=decisions,rows=rows,adverse_flags=adverse,
                allocation_gates_sha256=native._sha(HERE/'allocation-gates.json'),
                scope='Planning/commit/publication allocation vectors only; no instrumented latency comparison')

if __name__ == '__main__':
    result=analyze()
    (HERE/'allocation-analysis.json').write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
    print('Planning allocation gate:',result['planning_gate_passed'],'adverse:',len(result['adverse_flags']))
