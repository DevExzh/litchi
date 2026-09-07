#!/usr/bin/env python3
"""Validate pacing fields, then replay the unchanged provider lifecycle oracle."""
import argparse,copy,importlib.util,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('base_oracle',ROOT/'base-verify-report.py')
base=importlib.util.module_from_spec(spec);spec.loader.exec_module(base)
FIELDS=('transfer_paced_calls','transfer_delay_ns')
def require(ok,message):
    if not ok:raise ValueError(message)
def uint(value):
    require(type(value) is int and 0<=value<2**64,'expected u64 pacing counter');return value
def bounds(point,rate,cap):
    calls,delay=(uint(point[k]) for k in FIELDS);returned=uint(point['returned_bytes'])
    require(calls<=uint(point['logical_calls']),'paced calls exceed reads')
    if rate is None:require(calls==delay==0,'unconfigured pacing fabricates delay');return
    require((returned==0)==(calls==0)==(delay==0),'pacing zero state differs from returned bytes')
    require(calls<=returned<=calls*cap,'paced chunk counts do not bound returned bytes')
    lower=(returned*10**9+rate-1)//rate
    require(lower<=delay<=lower+max(0,calls-1),'sum of per-read ceiling delays is inconsistent')
ALLOCATION_FIELDS=('open_source_allocation_metrics','open_destination_allocation_metrics','plan_allocation_metrics','publication_allocation_metrics')
ALLOCATION_NUMBERS=('allocation_calls','deallocation_calls','reallocation_calls','failed_allocation_calls','allocated_bytes','deallocated_bytes','live_bytes_before','live_bytes_after','peak_live_bytes_before','peak_live_bytes_after','region_peak_live_bytes')
def allocation(report):
    mode=report.pop('instrumentation');allocator=report.pop('allocator');revision=report.pop('allocator_counter_revision',None)
    require(mode in ['none','system_allocator_operation_scoped'],'instrumentation identity')
    measured=mode!='none'
    require(allocator==('CountingSystemAllocator(std::alloc::System)' if measured else 'Rust system allocator'),'allocator identity')
    require(revision==('serialized_region_peak_v3' if measured else None),'allocator counter revision')
    scopes=json.loads((ROOT/'allocation-scopes.json').read_text())
    for name,expected in scopes.items():require(report.pop(name)==expected,'allocation scope '+name)
    for row in report['samples_raw']:
        timing=row['timings']
        for key in ALLOCATION_FIELDS:
            if not measured:
                require(key not in timing,'ordinary binary publishes allocator sample');continue
            v=timing.pop(key);require(set(v)=={'status','scope',*ALLOCATION_NUMBERS},'complete allocation sample')
            require(v['status']=='measured' and v['scope']=='operation_global_system_allocator','valid allocation observation')
            for name in ALLOCATION_NUMBERS:uint(v[name])
            require(v['failed_allocation_calls']==0,'allocation failure')
            require(v['reallocation_calls']<=v['allocation_calls'],'reallocation calls')
            require(v['live_bytes_before']+v['allocated_bytes']==v['live_bytes_after']+v['deallocated_bytes'],'allocator live-byte conservation')
            require(max(v['live_bytes_before'],v['live_bytes_after'])<=v['region_peak_live_bytes']<=v['peak_live_bytes_after'],'region peak bounds')
            require(v['peak_live_bytes_before']<=v['peak_live_bytes_after'] and v['live_bytes_before']<=v['peak_live_bytes_before'],'absolute peak bounds')
    return mode

def check_report(value):
    report=copy.deepcopy(value);require(report['schema']=='pptx_provider_lifecycle_v1','provider schema')
    instrumentation=allocation(report)
    config=report['provider_config'];require('transfer_bytes_per_second' in config,'missing transfer configuration')
    policy=config.pop('transfer_delay_policy')
    require(policy in ['separate-sleeps','minimum-service'],'unknown transfer policy')
    rate=config.pop('transfer_bytes_per_second')
    require(policy=='separate-sleeps' or rate is not None,'minimum service requires a rate')
    if rate is not None:
        uint(rate);require(1048576<=rate<=1099511627776 and report['provider']=='range','invalid transfer rate/provider')
    cap=config['max_range_bytes']
    for row in report['samples_raw']:
        previous={'source_reads':None,'destination_reads':None}
        for phase in row['phases']:
            delay=0
            for field in previous:
                point=phase[field]
                require(all(k in point for k in FIELDS),'missing pacing snapshot')
                if point['availability']=='unavailable':
                    require(all(point[k] is None for k in FIELDS),'unavailable pacing has values')
                    for k in FIELDS:point.pop(k)
                    previous[field]=None;continue
                delta=point['delta'];require(type(delta) is dict and all(k in delta for k in FIELDS),'missing pacing delta')
                bounds(point,rate,cap);bounds(delta,rate,cap)
                for k in FIELDS:
                    before=0 if previous[field] is None else previous[field][k]
                    require(point[k]>=before and delta[k]==point[k]-before,'pacing delta differs from snapshots')
                previous[field]={k:point[k] for k in FIELDS}
                delay+=delta['transfer_delay_ns']+delta['delayed_calls']*(config['delay_us'] or 0)*1000
                for k in FIELDS:point.pop(k);delta.pop(k)
            clock={'opened':'open_ns','planned':'plan_ns','published':'publication_ns'}.get(phase['label'])
            if clock:require(row['timings'][clock]>=delay,'combined nominal service floor exceeds enclosing serial API clock')
    result=base.check_report(report)
    return {'status':'pass','instrumentation':instrumentation,'transfer_bytes_per_second':rate,'transfer_delay_policy':policy,'base':result}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('report',type=Path);a=p.parse_args()
    print(json.dumps(check_report(base.load_json(a.report)),sort_keys=True))
