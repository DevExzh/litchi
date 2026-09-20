#!/usr/bin/env python3
"""Inject malformed evidence in memory; retain every captured byte unchanged."""
import copy,importlib.util,json
from pathlib import Path
from unittest.mock import patch
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('refusals0718',P/'analyze.py');A=importlib.util.module_from_spec(spec);spec.loader.exec_module(A)
def main():
    dest=P/'negative-checks.json';assert not dest.exists()
    original=A.analyze();assert original==A.read(P/'analysis.json')
    B=A.module(P/'brk-analysis.py','brk0718negative');brk_original=B.analyze();assert brk_original==A.read(P/'brk-analysis.json')
    protected={str(f.relative_to(P)):A.sha(f) for f in P.rglob('*') if f.is_file()}
    name='mapping-r1-generated';allocation='allocation-r1-generated.dhat.json'
    cases=[
        ('elapsed-mean',name+'.json',lambda v:v['results'][0]['elapsed_ns'].__setitem__('mean',v['results'][0]['elapsed_ns']['mean']+1)),
        ('command-cpu',name+'.receipt.json',lambda v:v['command'].__setitem__(2,'13')),
        ('decoded-target-size','mapping-r1-numbered-list.json',lambda v:v['results'][0]['corpus'].__setitem__('target_payload_bytes',1)),
        ('instrumentation',name+'.json',lambda v:v['tool'].__setitem__('instrumentation','none')),
        ('process-alignment',name+'.json',lambda v:v['results'][0]['source']['ordinary_save']['process_probe']['sample_deltas'][0].__setitem__('minor_faults',999999)),
        ('control-count',name+'.json',lambda v:v['results'][0]['source']['ordinary_save']['process_probe']['empty_adjacent_snapshot_controls'].pop()),
        ('negative-allocation',allocation,lambda v:v['pps'][0].__setitem__('tb',-1)),
        ('allocation-frame',allocation,lambda v:v['pps'][0]['fs'].__setitem__(0,99999999)),
    ]
    cases.append(('symbol-offset','symbolization.json',lambda v:next(iter(v['frames'].values())).__setitem__('virtual_address',next(iter(v['frames'].values()))['virtual_address']+1)))
    cases=[(label,filename,mutate,False,A.analyze) for label,filename,mutate in cases]
    cases += [
        ('marker-triplet',name+'.strace',lambda text:text.replace('"/proc/self/io"','"/proc/self/status"',1),True,A.analyze),
        ('brk-request-failure',name+'.strace',lambda text:text.replace('brk(NULL)','brk(0x1)',1),True,B.analyze),
    ]
    real_read=Path.read_text;checks=[]
    for label,filename,mutate,is_text,analyze in cases:
        def changed(path,*args,**kwargs):
            raw=real_read(path,*args,**kwargs)
            if path.resolve()==(P/filename).resolve():
                if is_text:return mutate(raw)
                value=json.loads(raw);mutate(value);return json.dumps(value)
            return raw
        with patch.object(Path,'read_text',changed):
            try:analyze()
            except (AssertionError,RuntimeError,ValueError) as error:checks.append(dict(name=label,rejected=True,error=str(error)))
            else:raise AssertionError(label+' was accepted')
    assert all(A.sha(P/n)==h for n,h in protected.items())
    assert A.analyze()==original and B.analyze()==brk_original
    dest.write_text(json.dumps(dict(status='pass',checks=checks,retained_inputs_unchanged=True,positive_exact_replay=True,analysis_sha256=A.sha(P/'analysis.json'),analyzer_sha256=A.sha(P/'analyze.py'),attribution_sha256=A.sha(P/'attribution.py'),brk_analysis_sha256=A.sha(P/'brk-analysis.json'),brk_analyzer_sha256=A.sha(P/'brk-analysis.py')),indent=2)+'\n');print('PASS eleven corruption refusals and exact replay')
if __name__=='__main__':main()
