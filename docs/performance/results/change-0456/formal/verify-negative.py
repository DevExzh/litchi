#!/usr/bin/env python3
"""Verify that corrupted sink observations are rejected without editing evidence."""
import copy,hashlib,importlib.util,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('oracle',ROOT/'verify-report.py');oracle=importlib.util.module_from_spec(spec);spec.loader.exec_module(oracle)
source=ROOT/'runs/5/report.json';original=json.loads(source.read_text());oracle.check_report(original)
def accepted(report):report['samples_raw'][0]['publication_sink']['accepted_bytes']+=1
def histogram(report):report['samples_raw'][0]['publication_sink']['write_size_buckets']['bytes_1_to_512']+=1
def cap(report):report['samples_raw'][0]['publication_sink']['largest_write']=65537
def missing(report):del report['samples_raw'][0]['publication_sink']
def output(report):report['samples_raw'][0]['output_sha256']='0'*64
results=[]
for mutation in [accepted,histogram,cap,missing,output]:
    report=copy.deepcopy(original);mutation(report)
    try:oracle.check_report(report)
    except (ValueError,KeyError,AssertionError) as error:results.append({'mutation':mutation.__name__,'status':'rejected','error':str(error)})
    else:raise AssertionError('mutation accepted: '+mutation.__name__)
value={'status':'pass','report_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'oracle_sha256':hashlib.sha256((ROOT/'verify-report.py').read_bytes()).hexdigest(),'results':results}
with (ROOT/'negative-results.json').open('x') as f:f.write(json.dumps(value,indent=2)+'\n')
print(json.dumps(value))
