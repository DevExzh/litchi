#!/usr/bin/env python3
"""Freeze independent timing and allocator matrices before their pilots."""
import datetime,hashlib,json,platform,subprocess
from pathlib import Path
ROOT=Path(__file__).resolve().parent
assert not (ROOT/'protocol.json').exists()
for tag in ['baseline-build-r3','candidate-build','final-pptx-r2','final-harness-r2','final-strict-r2']:
 assert json.loads((ROOT/'checks'/f'{tag}.json').read_text())['status']=='pass'
normal=[dict(provider=p,corpus=c,build=b,instrumentation='normal') for p in ['bytes','range'] for b in ['baseline','candidate'] for c in ['plain','media-rich']]
alloc=[dict(provider='bytes',corpus=c,build=b,instrumentation='alloc') for b in ['baseline','candidate'] for c in ['plain','media-rich']]
order=[dict(x,repeat='R1') for x in normal]+[dict(x,repeat='R2') for x in reversed(normal)]+[dict(x,repeat='R1') for x in alloc]+[dict(x,repeat='R2') for x in reversed(alloc)]
names=['baseline-build.json','candidate-build.json','capture.py','verify-report.py','base-verify-report.py','allocation-scopes.json']
v={'change':453,'status':'frozen','frozen_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'cpu':2,'workers':1,'samples':30,'warmups':3,'order':order,'reports':24,'normal_reports':16,'normal_samples':480,'allocator_reports':8,'allocator_samples':240,'pilot_lanes':list(range(8))+list(range(16,20)),'acceptance':{'media_plan_allocated_bytes_reduction_minimum':16773120,'media_plan_live_growth_reduction_minimum':16773120,'both_repeats':True,'absolute_regression_review_percent':5},'claims':'Normal API latency separate from instrumented allocator diagnostics. Simulated range 64KiB/200us/25MiB per second separate sleeps. No cold/native/scaling claim. Keep all phase and repeat flags.','corpora':json.loads((ROOT.parent/'change-0452/protocol.json').read_text())['corpora'],'bound_files':{n:hashlib.sha256((ROOT/n).read_bytes()).hexdigest() for n in names}}
(ROOT/'protocol.json').write_text(json.dumps(v,indent=2)+'\n')
(ROOT/'machine.json').write_text(json.dumps({'uname':list(platform.uname()),'lscpu':subprocess.check_output(['lscpu'],text=True),'rustc':subprocess.check_output(['rustc','+1.98.1','-Vv'],text=True),'meminfo':Path('/proc/meminfo').read_text()},indent=2)+'\n');print('FROZEN')
