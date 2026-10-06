#!/usr/bin/env python3
"""Post-capture exploratory association, not a causal or qualification test.

Reuses the prior 0939 / prospective 0940 100000 ns preparation threshold.
Stratifying fault counts by that threshold was selected after capture.
"""
import argparse
import hashlib
import json
from pathlib import Path
import statistics


def describe(values):
    return {'count':len(values),'median':statistics.median(values) if values else None,
            'min':min(values) if values else None,'max':max(values) if values else None}


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('analysis',type=Path)
    ap.add_argument('packet',type=Path)
    ap.add_argument('output',type=Path)
    args=ap.parse_args()
    data=args.analysis.read_bytes(); analysis=json.loads(data)
    groups=[]
    for arm in ('normal_on','allocator_on'):
        for route in ('owned_vec','owned_reader'):
            strata={'below_100000ns':[],'at_or_above_100000ns':[]}
            child_rows=[]
            for child in analysis['children']:
                if (child['arm'],child['route'],child['case'])!=(arm,route,'workload'):
                    continue
                descriptor=child['report']; raw=(args.packet/descriptor['path']).read_bytes()
                assert len(raw)==descriptor['bytes'] and hashlib.sha256(raw).hexdigest()==descriptor['sha256']
                report=json.loads(raw)
                assert len(report['preparation_ns'])==len(report['process']['prep'])==50
                counts={key:0 for key in strata}
                unavailable=0
                # Both evidence vectors share the same operation rank order.
                for elapsed,observation in zip(report['preparation_ns'],report['process']['prep']):
                    key='at_or_above_100000ns' if elapsed>=100000 else 'below_100000ns'
                    counts[key]+=1
                    if observation['delta'] is None:
                        unavailable+=1
                        continue
                    strata[key].append((elapsed,observation['delta']['minor_faults']))
                child_rows.append({'id':child['id'],'counts':counts,'unavailable_process_samples':unavailable})
            assert len(child_rows)==12
            groups.append({'arm':arm,'route':route,'children':child_rows,'strata':{
                key:{'preparation_ns':describe([v[0] for v in rows]),
                     'minor_faults':describe([v[1] for v in rows])}
                for key,rows in strata.items()}})
    result={'schema':'litchi.performance.0940.exploratory-preparation-fault-association.v1',
            'status':'descriptive_only','analysis':{'path':str(args.analysis),'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()},
            'threshold_ns':100000,'selection':'Threshold inherited from0939; fault stratification selected after0940capture.',
            'limits':['Pooled samples are not independent child replicates; no significance test.',
                      'Process preparation interval includes procfs observation work.',
                      'Association does not establish a causal allocator or OS mechanism.',
                      'No calibration subtraction, gate relaxation, or optimization adoption.'],
            'groups':groups}
    with args.output.open('x') as f:
        json.dump(result,f,indent=2,sort_keys=True);f.write('\n')


if __name__=='__main__':main()
