#!/usr/bin/env python3
"""Reject altered pilot evidence in memory and retain exact positive replay."""
import copy,importlib.util,json
from pathlib import Path
P=Path(__file__).resolve().parent
s=importlib.util.spec_from_file_location('negative_pilot0715',P/'analyze_pilot.py');A=importlib.util.module_from_spec(s);s.loader.exec_module(A)
def main():
    dest=P/'pilot-negative-checks.json';assert not dest.exists()
    encoded=(json.dumps(A.analyze(),indent=2)+'\n').encode();assert encoded==(P/'pilot-analysis.json').read_bytes()
    protected={f.name:A.sha(f) for f in P.iterdir() if f.is_file()};checks=[]
    cases=[
        ('elapsed-mean','candidate_B1-native-generated-counting_publish.json',lambda v:v['results'][0]['elapsed_ns'].__setitem__('mean',v['results'][0]['elapsed_ns']['mean']+1)),
        ('command-cpu','baseline_A1-native-generated-counting_publish.receipt.json',lambda v:v['command'].__setitem__(2,'13')),
        ('decoded-target-size','candidate_B1-native-numbered-list-counting_publish.json',lambda v:v['results'][0]['corpus'].__setitem__('target_payload_bytes',1)),
        ('failed-allocation','candidate_B1-allocator-generated-counting_publish.json',lambda v:v['results'][0]['operation_metrics']['allocation']['failed_allocation_calls']['values'].__setitem__(0,1)),
    ]
    read=A.load
    for name,filename,mutate in cases:
        def altered(path):
            value=read(path)
            if path==P/filename:value=copy.deepcopy(value);mutate(value)
            return value
        A.load=altered
        try:A.analyze()
        except (AssertionError,RuntimeError,ValueError) as error:checks.append(dict(name=name,rejected=True,error=str(error)))
        else:raise AssertionError(name)
        finally:A.load=read
    assert all(A.sha(P/n)==h for n,h in protected.items())
    dest.write_text(json.dumps(dict(status='pass',checks=checks,retained_inputs_unchanged=True,positive_exact_replay=True,analysis_sha256=A.sha(P/'pilot-analysis.json'),analyzer_sha256=A.sha(P/'analyze_pilot.py')),indent=2)+'\n')
    print('PASS four pilot corruptions rejected; retained inputs unchanged')
if __name__=='__main__':main()
