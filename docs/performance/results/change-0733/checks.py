#!/usr/bin/env python3
"""Negative controls against the actual profile parser and custody checks."""
import copy
import contextlib
import hashlib
import importlib.util
import io
import json
import shutil
import sys
import tempfile
from pathlib import Path
sys.dont_write_bytecode = True
P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('finish_analysis', P/'analyze.py')
a = importlib.util.module_from_spec(spec); spec.loader.exec_module(a)

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
checks = []
def reject(name, expected, action):
    try:
        action()
    except a.Invalid as error:
        message = str(error)
        assert expected in message, (name, message)
        checks.append(dict(name=name, rejected=True, reason=expected))
    else:
        raise AssertionError('accepted control: '+name)

with contextlib.redirect_stdout(io.StringIO()):
    a.main()
raw = (a.OLD/'captures/callgrind-0.callgrind').read_text()
profile = a.parse(a.OLD/'captures/callgrind-0.callgrind')
with tempfile.TemporaryDirectory(prefix='litchi-0733-checks-') as temporary:
    t = Path(temporary)
    def parse_mutated(value):
        path=t/'mutated.callgrind';path.write_text(value);return a.parse(path)
    reject('wrong event dimension', 'unexpected events', lambda:parse_mutated(raw.replace('events: Ir','events: Dr',1)))
    reject('wrong positions dimension', 'unexpected positions', lambda:parse_mutated(raw.replace('positions: line','positions: instr',1)))
    reject('wrong summary', 'profile totals disagree', lambda:parse_mutated(raw.replace('summary: 59578167','summary: 59578168',1)))
    reject('wrong self total', 'self sum differs', lambda:parse_mutated(raw.replace('summary: 59578167','summary: 59578168',1).replace('totals: 59578167','totals: 59578168',1)))
    reject('missing totals', 'profile totals disagree', lambda:parse_mutated(raw.replace('totals: 59578167','',1)))
    reject('dangling call association', 'incomplete profile', lambda:parse_mutated(raw+'\ncalls=1 0\n'))
    bad=copy.deepcopy(profile);bad['edges'][('foreign caller',a.PACKAGE)]=1
    reject('shared context cannot be partitioned', 'ambiguous positive-cost caller',lambda:a.partition(bad,a.PACKAGE,a.FINISH))
    bad2=copy.deepcopy(profile);bad2['edges'][(a.FINISH,a.PACKAGE)]+=1
    reject('nonconserving inclusive edge', 'incoming/outgoing costs disagree',lambda:a.partition(bad2,a.PACKAGE,a.FINISH))
    isolated=t/'packet';shutil.copytree(P,isolated,ignore=shutil.ignore_patterns('__pycache__'))
    a.P=isolated
    ancestry=json.loads((isolated/'ancestry.json').read_text());ancestry['change-0731']['sha256']='0'*64
    (isolated/'ancestry.json').write_text(json.dumps(ancestry))
    reject('altered ancestor binding','ancestor seal changed',a.custody)
    a.P=P
    isolated2=t/'witness-packet';shutil.copytree(P,isolated2,ignore=shutil.ignore_patterns('__pycache__'))
    a.P=isolated2
    plan=json.loads((isolated2/'witness-plan.json').read_text());plan['source_sha256']='0'*64
    (isolated2/'witness-plan.json').write_text(json.dumps(plan))
    reject('altered witness source receipt','witness source changed',a.witness)
    a.P=P
receipt=dict(status='passed',baseline_accepted=True,rejection_count=len(checks),checks=checks,
             analyzer_sha256=sha(P/'analyze.py'),script_sha256=sha(Path(__file__)))
(P/'checks.json').write_text(json.dumps(receipt,indent=2)+'\n')
print(f'PASS actual analyzer baseline and {len(checks)} parser/context/custody rejection controls')
