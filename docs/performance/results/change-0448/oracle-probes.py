#!/usr/bin/env python3
"""Reject corrupted pacing and inherited lifecycle evidence before freezing."""
import copy,importlib.util,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('oracle',ROOT/'verify-report.py');oracle=importlib.util.module_from_spec(spec);spec.loader.exec_module(oracle)
report=oracle.base.load_json(ROOT/'pilots/3/control/report.json');oracle.check_report(report)
point=lambda v:v['samples_raw'][0]['phases'][1]['source_reads']
mutations={
 'missing-policy':lambda v:v['provider_config'].pop('transfer_delay_policy'),
 'unknown-policy':lambda v:v['provider_config'].__setitem__('transfer_delay_policy','unknown'),
 'impossible-service-floor':lambda v:v['samples_raw'][0]['timings'].__setitem__('open_ns',0),
 'missing-rate':lambda v:v['provider_config'].pop('transfer_bytes_per_second'),
 'zero-rate':lambda v:v['provider_config'].__setitem__('transfer_bytes_per_second',0),
 'boolean-rate':lambda v:v['provider_config'].__setitem__('transfer_bytes_per_second',True),
 'rate-on-bytes':lambda v:v.__setitem__('provider','bytes'),
 'missing-counter':lambda v:point(v).pop('transfer_paced_calls'),
 'negative-delay':lambda v:point(v).__setitem__('transfer_delay_ns',-1),
 'oversized-delay':lambda v:point(v).__setitem__('transfer_delay_ns',2**64),
 'boolean-counter':lambda v:point(v).__setitem__('transfer_paced_calls',True),
 'invented-unavailable':lambda v:v['samples_raw'][0]['phases'][0]['source_reads'].__setitem__('transfer_delay_ns',1),
 'zero-transfer':lambda v:point(v).__setitem__('transfer_delay_ns',0),
 'bad-ceiling-sum':lambda v:point(v).__setitem__('transfer_delay_ns',point(v)['transfer_delay_ns']+point(v)['logical_calls']+1),
 'bad-delta':lambda v:point(v)['delta'].__setitem__('transfer_delay_ns',point(v)['delta']['transfer_delay_ns']+1),
 'counter-underflow':lambda v:point(v)['delta'].__setitem__('transfer_paced_calls',point(v)['delta']['transfer_paced_calls']+1),
 'missing-budget':lambda v:v['samples_raw'][0]['phases'][1]['source_budget'].pop('memory_used'),
 'wrong-output':lambda v:v['samples_raw'][0].__setitem__('output_sha256','0'*64),
 'inconsistent-timer':lambda v:v['samples_raw'][0]['timings'].__setitem__('api_sum_ns',1),
 'invented-pacing-control':lambda v:v['provider_config'].__setitem__('transfer_bytes_per_second',None),
}
rows=[]
for name,mutate in mutations.items():
    value=copy.deepcopy(report);mutate(value)
    try:oracle.check_report(value)
    except Exception as error:rows.append({'mutation':name,'rejected':str(error)})
    else:raise AssertionError('accepted '+name)
(ROOT/'oracle-probes.json').write_text(json.dumps({'change':448,'status':'pass','rejected':rows,'scope':'new pacing arithmetic/availability/CLI configuration plus inherited output, timer and budget report contracts; does not independently replay semantic publication'},indent=2)+'\n')
print(json.dumps({'status':'pass','rejected':len(rows)}))
