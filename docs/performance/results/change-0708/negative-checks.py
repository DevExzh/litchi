#!/usr/bin/env python3
"""Exercise verifier refusals on in-memory copies; never mutate raw evidence."""
import copy
import importlib.util
import json
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('analyzer0708',P/'analyze.py')
A=importlib.util.module_from_spec(spec);spec.loader.exec_module(A)
def main():
    plan=A.load_plan();phase='baseline-A1';job=A.expected_jobs(plan,phase)[0]
    raw=A.read(P/(job['name']+'.json'))['results'][0]
    receipt=A.read(P/(job['name']+'.receipt.json'))
    build=A.build_info('baseline','native')[0]
    A.validate_corpus(raw['corpus'],job['shape'],'positive')
    A.validate_source(raw,job,allocator=False)
    A.validate_receipt(receipt,job,phase,'baseline','native',plan,build,build['expected_source'])
    rows=[]
    def rejects(name,fn):
        try:fn()
        except (AssertionError,ValueError,RuntimeError,SystemExit) as error:
            rows.append(dict(name=name,rejected=True,message=str(error)));return
        raise AssertionError(name+' accepted invalid evidence')
    for key,value in [('shape','wrong'),('generator','wrong'),('name','wrong')]:
        bad=copy.deepcopy(raw['corpus']);bad[key]=value
        rejects('corpus-'+key,lambda:A.validate_corpus(bad,job['shape'],'negative'))
    for key,value in [('timing_scope','wrong'),('update_count',-1),('selected_worksheet_count',999)]:
        bad=copy.deepcopy(raw);bad['source']['xlsx_cell_values'][key]=value
        rejects('source-'+key,lambda:A.validate_source(bad,job,allocator=False))
    bad=copy.deepcopy(receipt);bad.pop('build_record_sha256')
    rejects('missing-build-binding',lambda:A.validate_receipt(bad,job,phase,'baseline','native',plan,build,build['expected_source']))
    result=dict(status='pass',analyzer_sha256=A.sha(P/'analyze.py'),checks=rows)
    (P/'negative-checks.json').write_text(json.dumps(result,indent=2)+'\n')
    print('PASS',len(rows),'in-memory negative verifier checks')
if __name__=='__main__':main()
