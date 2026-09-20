#!/usr/bin/env python3
"""Exercise semantic evidence rejection without rewriting any retained input."""
import copy
import importlib.util
import json
from pathlib import Path

P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('negative0717',P/'analyze.py')
A=importlib.util.module_from_spec(spec);spec.loader.exec_module(A)

def main():
    dest=P/'negative-checks.json'
    assert not dest.exists()
    original=A.analyze()
    assert original==json.loads((P/'analysis.json').read_text())
    protected={str(f.relative_to(P)):A.sha(f) for f in P.rglob('*') if f.is_file()}
    prefix='native-b1-generated'
    checks=[]
    cases=[
        ('elapsed-mean',prefix+'.json',lambda v:v['results'][0]['elapsed_ns'].__setitem__('mean',v['results'][0]['elapsed_ns']['mean']+1)),
        ('sample-permutation',prefix+'.json',lambda v:v['results'][0]['elapsed_ns']['sample_order'].__setitem__(0,v['results'][0]['elapsed_ns']['sample_order'][1])),
        ('command-cpu',prefix+'.receipt.json',lambda v:v['command'].__setitem__(2,'13')),
        ('decoded-target-size','native-b1-numbered-list.json',lambda v:v['results'][0]['corpus'].__setitem__('target_payload_bytes',1)),
        ('negative-process-counter',prefix+'.context.json',lambda v:v['wait4'].__setitem__('ru_nivcsw',-1)),
    ]
    probe='procfs-b1-generated.json'
    cases += [
        ('instrumentation',probe,lambda v:v['tool'].__setitem__('instrumentation','none')),
        ('process-alignment',probe,lambda v:v['results'][0]['source']['ordinary_save']['process_probe']['sample_deltas'][0].__setitem__('minor_faults',999999)),
        ('control-count',probe,lambda v:v['results'][0]['source']['ordinary_save']['process_probe']['empty_adjacent_snapshot_controls'].pop()),
        ('process-scope',probe,lambda v:v['results'][0]['source']['ordinary_save']['process_probe'].__setitem__('scope','owner_exclusive')),
    ]
    read=A.read
    for name,filename,mutate in cases:
        def altered(path):
            value=read(path)
            if Path(path)==P/filename:
                value=copy.deepcopy(value);mutate(value)
            return value
        A.read=altered
        try:
            A.analyze()
        except (AssertionError,RuntimeError,ValueError) as error:
            checks.append(dict(name=name,rejected=True,error=str(error)))
        else:
            raise AssertionError(name+' was accepted')
        finally:
            A.read=read
    assert all(A.sha(P/n)==h for n,h in protected.items())
    assert A.analyze()==original
    dest.write_text(json.dumps(dict(status='pass',checks=checks,retained_inputs_unchanged=True,
                                   positive_exact_replay=True,analysis_sha256=A.sha(P/'analysis.json'),
                                   analyzer_sha256=A.sha(P/'analyze.py')),indent=2)+'\n')
    print('PASS nine in-memory evidence corruptions rejected; retained inputs unchanged')

if __name__=='__main__': main()
