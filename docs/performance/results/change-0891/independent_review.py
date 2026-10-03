from __future__ import annotations
import datetime
import hashlib
import json
import math
import pathlib
import random
import statistics
import re
import sys
from typing import Any

ROOT = pathlib.Path(__file__).resolve().parent
CAP = ROOT / 'capture'
PLAN_PATH = ROOT / 'plan.json'

class CheckError(Exception):
    pass

def req(cond: bool, msg: str):
    if not cond:
        raise CheckError(msg)

def read(path: pathlib.Path):
    with path.open(encoding='utf-8') as f:
        return json.load(f)

def sha(path: pathlib.Path) -> str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):
            h.update(b)
    return h.hexdigest()

def descriptor(path: pathlib.Path) -> dict[str,Any]:
    st=path.stat()
    return {'bytes':st.st_size,'sha256':sha(path),'mode':st.st_mode & 0o777,'executable':bool(st.st_mode & 0o100)}

def check_descriptor(rec: Any, path: pathlib.Path, what: str):
    req(isinstance(rec,dict), f'{what}: descriptor missing')
    req(path.is_file(), f'{what}: file missing {path}')
    actual=descriptor(path)
    for k,v in actual.items():
        if k in rec:
            req(rec[k]==v, f'{what}: {k} mismatch')
    return actual

def floor_midpoint(a:int,b:int)->int:
    return a//2+b//2+((a%2+b%2)//2)

def nearest_rank(v:list[int], p:int)->int:
    rank=(len(v)*p+99)//100
    return v[max(0,min(len(v)-1,rank-1))]

def raw_stats(v:list[int])->dict[str,int]:
    req(type(v) is list and v and all(type(x) is int and x>0 for x in v), 'invalid elapsed vector')
    req(v==sorted(v), 'elapsed vector is not sorted')
    return {'min':v[0], 'p50':floor_midpoint(v[(len(v)-1)//2],v[len(v)//2]), 'p95':nearest_rank(v,95), 'p99':nearest_rank(v,99), 'max':v[-1]}

def bootstrap(ratios:list[float], seed=891, draws=10000, low=249, high=9749)->dict[str,Any]:
    req(ratios, 'empty ratio vector')
    rng=random.Random(seed)
    vals=[]
    for _ in range(draws):
        vals.append(statistics.median([ratios[rng.randrange(len(ratios))] for _ in ratios]))
    ordered=sorted(vals)
    return {'seed':seed,'draws':draws,'low_index':low,'high_index':high,
            'median_ratio':statistics.median(ratios),'ci95_low':ordered[low],
            'ci95_high':ordered[high],'ratios':ratios}

def managed_item(plan, role, mode):
    x=plan['managed_binaries'][role][mode]
    if isinstance(x,dict) and isinstance(x.get('copy'),dict): x=x['copy']
    return x

def expected_rows(plan):
    cases=plan['cases']; rows=[]; ordinal=0
    for block in range(plan['native_blocks']):
        for ci,case in enumerate(cases):
            order=('before','after') if (block+ci)%2==0 else ('after','before')
            pair=f'native-b{block:02d}-c{ci:02d}'
            for role in order:
                rows.append({'ordinal':ordinal,'lane':'native','block':block,'case_index':ci,'case':case,'role':role,'pair':pair,'warmup':plan['native_warmup'],'samples':plan['native_samples']})
                ordinal+=1
    for block in range(plan['allocator_blocks']):
        order=('before','after') if block%2==0 else ('after','before')
        pair=f'allocator-b{block:02d}-c00'
        for role in order:
            rows.append({'ordinal':ordinal,'lane':'allocator','block':block,'case_index':0,'case':cases[0],'role':role,'pair':pair,'warmup':plan['allocator_warmup'],'samples':plan['allocator_samples']})
            ordinal+=1
    return rows

def baseline_expected():
    out={}
    b=ROOT.parent/'change-0887'/'baseline-qualification'
    for case in plan['cases']:
        d=b/case
        rep=read(d/'report.json')
        out[case]={'corpus':rep['results'][0]['corpus'], 'report':rep,
                   'producer':read(d/'producer.json') if (d/'producer.json').is_file() else None}
    return out

def compare_report_semantics(rep, row, binary, expected, producer, mode):
    n=row['samples']; case=row['case']; role=row['role']
    req(rep.get('schema_version')==1, f'{row["ordinal"]}: schema version')
    tool=rep.get('tool',{})
    req(tool.get('name')=='litchi-perf-baseline' and tool.get('version')=='0.1.0', f'{row["ordinal"]}: tool')
    req(tool.get('profile')=='release' and tool.get('target_os')=='linux' and tool.get('target_arch')=='x86_64', f'{row["ordinal"]}: target')
    req(tool.get('binary')==('litchi-perf-baseline' if mode=='native' else 'litchi-perf-baseline-alloc'), f'{row["ordinal"]}: binary name')
    req(tool.get('instrumentation')==('none' if mode=='native' else 'system_allocator_operation_scoped'), f'{row["ordinal"]}: instrumentation')
    if mode=='allocator': req(tool.get('allocator_counter_revision')=='serialized_region_peak_v3', f'{row["ordinal"]}: counter revision')
    conf=rep.get('configuration',{})
    req(conf.get('samples_per_case')==n and conf.get('warmup_iterations_per_case')==row['warmup'], f'{row["ordinal"]}: sample protocol')
    req(conf.get('filesystem_cache_states')==['warm'] and conf.get('filesystem_fresh_child_per_sample') is True and conf.get('filesystem_process_isolated') is True, f'{row["ordinal"]}: filesystem protocol')
    req(conf.get('xlsx_shapes')==['medium'] and conf.get('cases')==[case], f'{row["ordinal"]}: case config')
    req(rep.get('environment',{}).get('cpu_affinity')=='12', f'{row["ordinal"]}: CPU affinity')
    ident=rep.get('binary_identity',{})
    req(ident.get('path')==binary['path'] and ident.get('binary_sha256')==binary['sha256'] and ident.get('binary_bytes')==binary['bytes'] and ident.get('mode_bits')==binary['mode'] and ident.get('executable') is True and ident.get('profile')=='release', f'{row["ordinal"]}: report binary identity')
    results=rep.get('results'); req(type(results) is list and len(results)==1 and results[0].get('case')==case, f'{row["ordinal"]}: result case')
    result=results[0]; elapsed=result.get('elapsed_ns',{}); raw=elapsed.get('samples'); stats=raw_stats(raw)
    req(len(raw)==n and elapsed.get('sample_order') is not None and sorted(elapsed['sample_order'])==list(range(n)), f'{row["ordinal"]}: raw/order')
    for k,v in stats.items(): req(elapsed.get(k)==v, f'{row["ordinal"]}: {k} mismatch')
    req(result.get('corpus')==expected['corpus'], f'{row["ordinal"]}: corpus mismatch')
    op=result.get('operation_metrics')
    if mode=='native' and case != plan['cases'][0]:
        req(op is None and 'source' not in result, f'{row["ordinal"]}: control unexpectedly has operation metrics')
    if mode=='native' and case == plan['cases'][0]:
        req(isinstance(op,dict), f'{row["ordinal"]}: primary native operation metrics')
    if mode=='allocator':
        req(isinstance(op,dict), f'{row["ordinal"]}: allocator operation metrics')
    if op is not None:
        req(op.get('sample_count')==n and op.get('sample_indices')==elapsed['sample_order'], f'{row["ordinal"]}: metric alignment')
        proc=op.get('process'); req(isinstance(proc,dict) and proc.get('status')=='measured', f'{row["ordinal"]}: process metrics')
        peak=proc.get('peak_rss_bytes')
        req(isinstance(peak,dict) and peak.get('status')=='measured', f'{row["ordinal"]}: process RSS metrics')
        pv=peak.get('values'); req(type(pv) is list and len(pv)==n and all(type(x) is int and x>0 for x in pv), f'{row["ordinal"]}: process RSS vector')
        if mode=='native': req(peak.get('scope')=='process_lifetime_high_water_after_not_operation_peak', f'{row["ordinal"]}: native RSS scope')
    expected_prod=expected['producer']
    if expected_prod is None:
        req(producer is None, f'{row["ordinal"]}: unexpected producer sidecar')
    else:
        req(producer==expected_prod, f'{row["ordinal"]}: producer sidecar differs from sealed fixture')
    if case==plan['cases'][0]:
        req(producer is not None, f'{row["ordinal"]}: primary producer sidecar missing')
        source=result.get('source',{}); scan=source.get('producer_full_cell_scan'); req(isinstance(scan,dict), f'{row["ordinal"]}: producer scan missing')
        base_scan=expected['report']['results'][0]['source']['producer_full_cell_scan']
        dynamic={'sample_indices','observed_cell_counts','observed_numeric_counts','observed_text_counts','observed_digest_sha256','source_version_stable'}
        for k,v in base_scan.items():
            if k not in dynamic: req(scan.get(k)==v, f'{row["ordinal"]}: scan {k} mismatch')
        req(scan['sample_indices']==elapsed['sample_order'], f'{row["ordinal"]}: scan sample alignment')
        req(scan['observed_cell_counts']==[scan['expected_cell_count']]*n, f'{row["ordinal"]}: cell counts')
        req(scan['observed_numeric_counts']==[scan['expected_numeric_count']]*n, f'{row["ordinal"]}: numeric counts')
        req(scan['observed_text_counts']==[scan['expected_text_count']]*n, f'{row["ordinal"]}: text counts')
        req(scan['observed_digest_sha256']==[scan['oracle_digest_sha256']]*n, f'{row["ordinal"]}: oracle digests')
        req(scan['source_version_stable']==[True]*n, f'{row["ordinal"]}: source version')
        req(scan['physical_cold_eligible'] is False and scan['runtime_route_observed'] is False, f'{row["ordinal"]}: route/cold claim')
    return {'stats':stats, 'sample_order':elapsed['sample_order'], 'result':result}

plan=read(PLAN_PATH)
req(plan.get('schema')=='litchi.xlsx-x14ac-experiment-0891.v1','plan schema')
req(plan.get('status')=='frozen' and plan.get('total_children')==72 and plan.get('bootstrap_seed')==891 and plan.get('bootstrap_draws')==10000 and plan.get('bootstrap_ci_indices')==[249,9749], 'plan protocol')
req(plan.get('cpu_affinity')=='12' and plan.get('cache')=='warm' and plan.get('profile')=='release' and plan.get('fresh_child') is True and plan.get('fail_fast') is True, 'plan execution')
plan_sha=sha(PLAN_PATH)
exp=read(CAP/'experiment.json')
req(exp.get('schema')=='litchi.xlsx-x14ac-experiment-receipt-0891.v1' and exp.get('status')=='pass' and exp.get('children')==72 and exp.get('native_children')==64 and exp.get('allocator_children')==8 and exp.get('no_retries_or_replacements') is True, 'experiment receipt')
req(exp.get('plan_sha256')==plan_sha, 'experiment plan hash')
rows=expected_rows(plan)
req(exp.get('rows')==[{'path':f'runs/{r["ordinal"]:03d}-{r["lane"]}-{r["pair"]}-{r["role"]}/receipt.json','row':r} for r in rows], 'frozen schedule')
expected=baseline_expected()

child={}
issues=[]
artifact_counts=collections = {'command':0,'report':0,'peak':0,'stdout':0,'stderr':0,'producer':0,'semantic':0}
previous_finished=None
for row in rows:
    path=CAP/f'runs/{row["ordinal"]:03d}-{row["lane"]}-{row["pair"]}-{row["role"]}'
    rp=path/'receipt.json'; req(rp.is_file(), f'missing receipt {row["ordinal"]}')
    rec=read(rp)
    req(rec.get('schema')=='litchi.xlsx-x14ac-child-receipt-0891.v1' and rec.get('row')==row and rec.get('exit_code')==0 and rec.get('plan_sha256')==plan_sha, f'receipt identity {row["ordinal"]}')
    mode='native' if row['lane']=='native' else 'allocator'
    m=managed_item(plan,row['role'],mode)
    binary={'path':m['path'],'bytes':m['bytes'],'sha256':m['sha256'],'mode':m['mode'],'executable':m['executable']}
    rb=rec.get('binary',{}); ri=rb.get('identity',{}) if isinstance(rb,dict) else {}
    req(rb.get('path')==binary['path'] and ri.get('sha256')==binary['sha256'] and ri.get('bytes')==binary['bytes'] and ri.get('mode')==binary['mode'] and ri.get('executable') is True, f'binary custody {row["ordinal"]}')
    binpath=pathlib.Path(binary['path'])
    if binpath.is_file():
        d=descriptor(binpath); req(d['sha256']==binary['sha256'] and d['bytes']==binary['bytes'] and d['mode']==binary['mode'] and d['executable'] is True, f'binary file mismatch {row["ordinal"]}')
    cmdp=path/'command.json'; cmd=read(cmdp); artifact_counts['command']+=1
    req(cmd.get('schema')=='litchi.xlsx-x14ac-child-command-0891.v1' and cmd.get('row')==row and cmd.get('environment')=={'PATH':'/usr/bin:/bin','LC_ALL':'C','LANG':'C','TZ':'UTC'}, f'command identity/env {row["ordinal"]}')
    req(cmd.get('argv')==rec.get('argv') and cmd.get('binary')==rec.get('binary') and cmd.get('cwd')==rec.get('cwd') and cmd.get('plan_sha256')==plan_sha and cmd.get('exit_code')==0, f'command receipt {row["ordinal"]}')
    reportp=path/'report.json'; peakp=path/'peak-rss-kib.txt'; stdoutp=path/'stdout.txt'; stderrp=path/'stderr.txt'
    for nm,pp,ck in [('report.json',reportp,'report'),('peak-rss-kib.txt',peakp,'peak'),('stdout.txt',stdoutp,'stdout'),('stderr.txt',stderrp,'stderr')]:
        check_descriptor(rec.get('artifacts',{}).get(nm),pp,f'{row["ordinal"]} {nm}'); artifact_counts[ck]+=1
    prodp=path/'producer.json'; prod=None
    if prodp.is_file():
        prod=read(prodp); check_descriptor(rec.get('artifacts',{}).get('producer.json'),prodp,f'{row["ordinal"]} producer'); artifact_counts['producer']+=1
    else:
        req('producer.json' not in rec.get('artifacts',{}), f'{row["ordinal"]}: receipt advertises absent producer')
    semanticp=path/'semantic.json'; req(semanticp.is_file(),f'{row["ordinal"]}: semantic receipt missing'); sem=read(semanticp); artifact_counts['semantic']+=1
    req(sem.get('status')=='pass' and sem.get('report_sha256')==sha(reportp) and sem.get('producer_sha256')==(sha(prodp) if prodp.is_file() else None),f'{row["ordinal"]}: semantic receipt')
    started=datetime.datetime.fromisoformat(rec['started_utc']); finished=datetime.datetime.fromisoformat(rec['finished_utc']); req(finished>=started,f'{row["ordinal"]}: interval');
    if previous_finished is not None: req(started>=previous_finished,f'{row["ordinal"]}: nonsequential')
    previous_finished=finished
    peak_text=peakp.read_text(encoding='utf-8').strip(); req(re.fullmatch(r'[0-9]+',peak_text) is not None,f'{row["ordinal"]}: peak text'); peak=int(peak_text); req(peak>0,f'{row["ordinal"]}: peak nonpositive')
    # Construct expected argv from the retained paths and exact output files.
    expected_argv=['/usr/bin/time','-f','%M','-o',str(peakp),'/usr/bin/taskset','-c','12',binary['path'],'--case',row['case'],'--xlsx-shape','medium','--warmup',str(row['warmup']),'--samples',str(row['samples']),'--filesystem-cache','warm','--json',str(reportp),'--producer-evidence',str(prodp)]
    req(rec.get('argv')==expected_argv,f'{row["ordinal"]}: argv')
    rep=read(reportp)
    checked=compare_report_semantics(rep,row,binary,expected[row['case']],prod,mode)
    child[row['ordinal']]={'row':row,'report':rep,'stats':checked['stats'],'peak_rss_kib':peak,'allocation':(rep['results'][0].get('operation_metrics',{}).get('allocation') if mode=='allocator' else None)}

req(len(child)==72 and previous_finished is not None,'child population')
# Pair summaries

def pair_groups(lane,case):
    sel=[v for v in child.values() if v['row']['lane']==lane and v['row']['case']==case]
    g={}
    for x in sel:
        g.setdefault(x['row']['pair'],{})[x['row']['role']]=x
    req(len(g)==(plan['native_blocks'] if lane=='native' else plan['allocator_blocks']),f'{lane}/{case}: pair count')
    out=[]
    for name in sorted(g):
        req(set(g[name])=={'before','after'},f'{name}: pair roles')
        out.append({'pair':name,'before':g[name]['before'],'after':g[name]['after']})
    return out

def metric_summary(pairs, metric):
    vals=[]
    for p in pairs:
        b=p['before']['stats'][metric] if metric in p['before']['stats'] else p['before']['peak_rss_kib']
        a=p['after']['stats'][metric] if metric in p['after']['stats'] else p['after']['peak_rss_kib']
        vals.append({'pair':p['pair'],'before':b,'after':a,'ratio':a/b})
    bs=bootstrap([x['ratio'] for x in vals])
    bs['pairs']=vals
    return bs

native={}
for case in plan['cases']:
    pairs=pair_groups('native',case)
    native[case]={k:metric_summary(pairs,k) for k in ('p50','p95','p99')}
    native[case]['peak_rss_kib']=metric_summary(pairs,'peak_rss_kib')

# Allocator counter summaries with raw vectors and medians.
alloc_pairs=pair_groups('allocator',plan['cases'][0])
alloc_metrics={}
failed_children=[]
for metric in ('allocation_calls','reallocation_calls','allocated_bytes','region_peak_live_bytes'):
    vals=[]
    for p in alloc_pairs:
        def av(item):
            alloc=item['allocation']; req(isinstance(alloc,dict) and alloc.get('status')=='measured',f'{p["pair"]}: allocation metrics')
            x=alloc.get(metric,{}); req(isinstance(x,dict) and x.get('status')=='measured' and x.get('scope')=='operation_global_system_allocator',f'{p["pair"]}: {metric} scope')
            v=x.get('values'); req(type(v) is list and len(v)==item['row']['samples'] and all(type(z) is int and z>=0 for z in v),f'{p["pair"]}: {metric} vector')
            return v
        bv,avv=av(p['before']),av(p['after'])
        b=statistics.median(bv); a=statistics.median(avv)
        vals.append({'pair':p['pair'],'before':b,'after':a,'ratio':a/b,'before_vector':bv,'after_vector':avv})
    s=bootstrap([x['ratio'] for x in vals]); s['pairs']=vals; alloc_metrics[metric]=s
for item in child.values():
    if item['row']['lane']=='allocator':
        alloc=item['allocation']; fv=alloc['failed_allocation_calls']['values']; req(fv==[0]*item['row']['samples'],f'{item["row"]["ordinal"]}: failed allocations')
        failed_children.append({'ordinal':item['row']['ordinal'],'pair':item['row']['pair'],'role':item['row']['role'],'values':fv,'all_zero':all(x==0 for x in fv),'max':max(fv)})

# All threshold flags, retaining each individual pair rather than only one summary per metric.
pair_threshold=1.05
primary_case=plan['cases'][0]
primary_flags=[]; control_flags=[]
for case,summary in native.items():
    target=primary_flags if case==primary_case else control_flags
    for metric,point in summary.items():
        for pair in point['pairs']:
            if pair['ratio']>pair_threshold:
                target.append({'case':case,'metric':metric,'pair':pair['pair'],'before':pair['before'],'after':pair['after'],'ratio':pair['ratio']})
spread=[]
for case in plan['cases']:
    for role in ('before','after'):
        vals=[x['stats']['p50'] for x in child.values() if x['row']['lane']=='native' and x['row']['case']==case and x['row']['role']==role]
        req(len(vals)==plan['native_blocks'],f'{case}/{role}: spread population')
        spread.append({'case':case,'role':role,'values':vals,'min':min(vals),'max':max(vals),'max_over_min':max(vals)/min(vals),'threshold':1.20,'flag':max(vals)/min(vals)>1.20})
primary=native[primary_case]['p50']
gate={'median_ratio_at_most':plan['primary_gate']['median_ratio_at_most'],'ci95_upper_below':plan['primary_gate']['ci95_upper_below'],'median_ratio':primary['median_ratio'],'ci95_high':primary['ci95_high'],'pass':primary['median_ratio']<=0.98 and primary['ci95_high']<1.0}

# Cross-check reduction if available, with exact independent result comparison for all generated summaries.
reduction_path=ROOT/'reduction.json'; reduction=read(reduction_path) if reduction_path.is_file() else None
reduction_match=None
if reduction is not None:
    req(reduction.get('plan_sha256')==plan_sha and reduction.get('experiment_sha256')==sha(CAP/'experiment.json'),'reduction bindings')
    reduction_match=True
    # Every numeric summary is compared at JSON-value granularity for fields independently produced.
    for case in plan['cases']:
        for metric, rmetric in (('p50','p50_ns'),('p95','p95_ns'),('p99','p99_ns')):
            lhs=native[case][metric]; rhs=reduction['native'][case][rmetric]
            req(lhs['ratios']==rhs['ratios'] and lhs['median_ratio']==rhs['median_ratio'] and lhs['ci95_low']==rhs['ci95_low'] and lhs['ci95_high']==rhs['ci95_high'],'reduction mismatch')
        lhs=native[case]['peak_rss_kib']; rhs=reduction['native'][case]['peak_rss_kib']
        req(lhs['ratios']==rhs['ratios'] and lhs['median_ratio']==rhs['median_ratio'] and lhs['ci95_low']==rhs['ci95_low'] and lhs['ci95_high']==rhs['ci95_high'],'reduction RSS mismatch')
    for metric in ('allocation_calls','allocated_bytes','region_peak_live_bytes'):
        lhs=alloc_metrics[metric]; rhs=reduction['allocator_primary'][metric]
        req(lhs['ratios']==rhs['ratios'] and lhs['median_ratio']==rhs['median_ratio'] and lhs['ci95_low']==rhs['ci95_low'] and lhs['ci95_high']==rhs['ci95_high'],'reduction allocator mismatch')
    req(reduction.get('primary_native_p50_gate',{}).get('pass')==gate['pass'],'reduction gate mismatch')

verification_path=ROOT/'verification.json'
verification_sha=sha(verification_path) if verification_path.is_file() else None
review={
 'schema':'litchi.xlsx-x14ac-capture-review-0891.v1',
 'status':'pass',
 'reviewer':'/root/baseline_review_0885',
 'scope':{'independent_raw_replay':True,'timed_workload_rerun':False,'cargo_rerun':False,'source_mutation':False},
 'replayed':{'plan_sha256':plan_sha,'capture_experiment_sha256':sha(CAP/'experiment.json'),'reduction_sha256':sha(reduction_path) if reduction_path.is_file() else None,'verification_sha256':verification_sha,'children':{'native':64,'allocator':8},'raw_sample_counts':{'native':64*50,'allocator':8*15},'artifact_hashes_checked':artifact_counts,'sequential_intervals_checked':72,'no_retries_or_replacements':exp['no_retries_or_replacements'],'reduction_exact_match':reduction_match,'statistics':'recomputed from each retained sorted raw elapsed vector using floored midpoint p50 and nearest-rank p95/p99; no sample pooling; RSS ratios use retained GNU /usr/bin/time peak-rss-kib.txt'},
 'native':native,
 'allocator_primary':alloc_metrics,
 'allocator_failed_allocation_calls':{'all_zero':all(x['all_zero'] for x in failed_children),'children':failed_children},
 'primary_native_p50_gate':gate,
 'primary_review_flags':primary_flags,
 'control_review_flags':control_flags,
 'native_p50_spread_flags':spread,
 'admissibility':{'semantic_population':{'native_children':64,'allocator_children':8,'native_cases':{c:16 for c in plan['cases']},'native_children_per_case_role':8,'allocator_case':primary_case,'allocator_children_per_role':4,'warmup_samples':{'native':{'warmup':3,'samples':50},'allocator':{'warmup':3,'samples':15}},'all_children_exit_code_zero':True,'producer_sidecars_checked':artifact_counts['producer']},'interpretation':{'primary_gate':'pass','primary_latency_evidence':'Primary native p50 whole-child paired median and bootstrap CI satisfy the frozen gate.','controls':'All individual >1.05 pair flags are retained below; control p99 tail flags remain review evidence and prevent treating the primary result as a universal latency claim.','rss':'GNU time process-lifetime high-water RSS is reported descriptively; it is not an adoption gate and no causal RSS mechanism is inferred.','allocator':'Allocation/reallocation/bytes/region-peak counters are separate observations; failed allocation calls are zero and no allocator adoption gate is applied.','scope':'Four fixed corpus cases, warm filesystem mode, CPU affinity 12 without exclusive-host or physical-cold-cache guarantees, release binaries, fresh sequential children; no general workload claim.'}},
 'issues':[],
 'uncertainty':['Control p99 flags are retained as observed whole-child tail variation; this replay cannot attribute mechanism or establish universal/exclusive-host behavior.','The capture and source semantics do not by themselves establish correctness outside the sealed fixed corpus and quality receipts.'],
 'performance_claim':None,
}
# Add concise check counts, not raw report blobs.
review['flags_summary']={'pair_ratio_threshold':1.05,'primary_count':len(primary_flags),'control_count':len(control_flags),'spread_threshold':1.20,'spread_flag_count':sum(x['flag'] for x in spread)}
out=pathlib.Path('/tmp/independent0891-review.json')
out.write_text(json.dumps(review,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('PASS independent replay; wrote',out)
print('plan',plan_sha,'capture',review['replayed']['capture_experiment_sha256'],'reduction',review['replayed']['reduction_sha256'],'verification',verification_sha)
print('flags',len(primary_flags),len(control_flags),'spread',sum(x['flag'] for x in spread),'alloc failed',len(failed_children))
print('gate',gate)
for c in plan['cases']:
  print(c,'p50',native[c]['p50']['median_ratio'],native[c]['p50']['ci95_low'],native[c]['p50']['ci95_high'],'p95',native[c]['p95']['median_ratio'],'p99',native[c]['p99']['median_ratio'],'rss',native[c]['peak_rss_kib']['median_ratio'])
for m,v in alloc_metrics.items(): print('alloc',m,v['median_ratio'],v['ci95_low'],v['ci95_high'],'pairs',[(x['before'],x['after']) for x in v['pairs']])
