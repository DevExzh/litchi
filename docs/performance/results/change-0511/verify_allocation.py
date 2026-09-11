"""Compare measured operation allocation deltas without instrumented timing claims."""
import json

def validate(here):
    fields=['allocation_calls','deallocation_calls','reallocation_calls','failed_allocation_calls','allocated_bytes','deallocated_bytes']
    cases={};observations={}
    for stage in ['before','after']:
        for repeat in ['r1','r2']:
            report=json.loads((here/stage/f'allocator-{repeat}-report.json').read_text())
            for row in report['results']:
                metrics=row['operation_metrics'];allocation=metrics['allocation'];assert allocation['status']=='measured'
                assert len(metrics['sample_indices'])==30 and sorted(metrics['sample_indices'])==list(range(30))
                data={}
                for field in fields:
                    value=allocation[field];assert value['status']=='measured' and len(value['values'])==30
                    assert len(set(value['values']))==1
                    data[field]=value['values'][0]
                peak=allocation['region_peak_live_bytes'];base=allocation['live_bytes_before']
                assert peak['status']=='measured' and base['status']=='measured'
                assert len(peak['values'])==len(base['values'])==30
                increments=[p-b for p,b in zip(peak['values'],base['values'])]
                assert min(increments)>=0
                data['incremental_region_peak_range_bytes']=[min(increments),max(increments)]
                observations[stage,repeat,row['case']]=data
                if row['case'] not in cases:cases[row['case']]=data
                else:assert all(cases[row['case']][f]==data[f] for f in fields)
    pairs=[]
    for repeat in ['r1','r2']:
        for case in sorted(cases):
            before,after=[observations[s,repeat,case] for s in ['before','after']]
            pairs.append({'repeat':repeat,'case':case,'before_incremental_region_peak_range_bytes':before['incremental_region_peak_range_bytes'],'after_incremental_region_peak_range_bytes':after['incremental_region_peak_range_bytes']})
    return {'samples':1080,'operation_allocation_deltas_identical_across_all_four_captures':True,'constant_operation_values':{case:{f:data[f] for f in fields} for case,data in cases.items()},'incremental_region_peaks':pairs,'scope':'Global allocator deltas bracket the existing operation; region peak minus starting live bytes is incremental operation demand, not standalone document/process peak; callbacks and temporary caller values follow existing timer scope','collection_note':'Captured from frozen binaries while all-feature test compilation ran; instrumented elapsed and RSS excluded from native comparisons','cfb_guard_allocation_metrics':'unavailable, not zero'}
