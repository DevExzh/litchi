#!/usr/bin/env python3
"""Reject in-memory corruptions and the actual stale preflight manifest."""
import copy
import importlib.util
import json
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('A0709',P/'analyze.py');A=importlib.util.module_from_spec(spec);spec.loader.exec_module(A)
def main():
    plan=A.load_plan();corpus=next(c for c in plan['corpora'] if c['id']=='numbered-list')
    job=next(j for j in A.expected_jobs(plan,'native') if j['corpus']['id']=='numbered-list' and j['phase']=='counting_publish')
    result=A.read(P/(job['name']+'.json'))['results'][0]
    A.result_corpus_identity(result,result['source']['ordinary_save'],corpus,'positive')
    rows=[]
    def rejects(name,fn):
        try:fn()
        except (AssertionError,RuntimeError,ValueError,KeyError) as error:
            rows.append(dict(name=name,rejected=True,message=str(error)));return
        raise AssertionError(name+' accepted invalid evidence')
    stale=A.read(P/'preflight/numbered-list.json')['results'][0]
    rejects('actual-stale-compressed-manifest',lambda:A.result_corpus_identity(stale,stale['source']['ordinary_save'],corpus,'stale'))
    for key,value in [('generator','wrong'),('target_payload_bytes',1242),('uncompressed_payload_bytes',1242),('target_payload_sha256','0'*64)]:
        bad=copy.deepcopy(result);bad['corpus'][key]=value
        rejects('manifest-'+key,lambda:A.result_corpus_identity(bad,bad['source']['ordinary_save'],corpus,'negative'))
    bad=copy.deepcopy(result['elapsed_ns']);bad['mean']+=1
    rejects('stale-mean',lambda:A.validate_elapsed(bad,100,'negative'))
    bad=copy.deepcopy(result);bad['source']['ordinary_save']['timing_scope']='wrong'
    rejects('wrong-timing-scope',lambda:A.validate_ordinary_save(bad,corpus,'counting_publish','native',100,'negative'))
    (P/'negative-checks.json').write_text(json.dumps(dict(status='pass',analyzer_sha256=A.sha(P/'analyze.py'),checks=rows),indent=2)+'\n');print('PASS',len(rows),'negative verifier checks')
if __name__=='__main__':main()
