#!/usr/bin/env python3
"""Exercise the final sink consistency guard on immutable in-memory report copies."""
import copy,hashlib,importlib.util,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('final_verifier',ROOT/'verify.py');v=importlib.util.module_from_spec(spec);spec.loader.exec_module(v)
report=json.loads((ROOT/'runs/5/report.json').read_text());v.check_sink_consistency(report)
results=[]
for largest in [1,2**64-1]:
    mutated=copy.deepcopy(report);mutated['samples_raw'][0]['publication_sink']['largest_write']=largest
    try:v.check_sink_consistency(mutated)
    except AssertionError as error:results.append({'largest_write':largest,'status':'rejected','error':str(error)})
    else:raise AssertionError('inconsistent maximum accepted')
record={'status':'pass','verifier_sha256':hashlib.sha256((ROOT/'verify.py').read_bytes()).hexdigest(),'results':results}
with (ROOT/'final-negative-results.json').open('x') as f:f.write(json.dumps(record,indent=2)+'\n')
print(json.dumps(record))
