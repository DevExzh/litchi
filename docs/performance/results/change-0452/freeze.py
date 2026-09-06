#!/usr/bin/env python3
"""Freeze matched source/binary/protocol identities before pilots or formal measurements."""
import datetime,hashlib,json,platform,subprocess
from pathlib import Path
ROOT=Path(__file__).resolve().parent
assert not (ROOT/'protocol.json').exists() and not (ROOT/'runs').exists()
for tag in ['baseline-harness-build','candidate-harness-build-r3','final-strict-r3','final-opc-tests-r3','final-pptx-tests-r3','final-harness-tests-r3']:
    assert json.loads((ROOT/'checks'/f'{tag}.json').read_text())['status']=='pass'
lanes=[{'provider':provider,'build':build,'corpus':corpus} for provider in ['bytes','range'] for build in ['baseline','candidate'] for corpus in ['plain','media-rich']]
order=[dict(lane,repeat='R1') for lane in lanes]+[dict(lane,repeat='R2') for lane in reversed(lanes)]
sha=lambda name:hashlib.sha256((ROOT/name).read_bytes()).hexdigest()
v={'change':452,'status':'frozen','frozen_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'cpu':2,'workers':1,'samples':30,'warmups':3,'order':order,'reports':16,'retained_samples':480,'acceptance':{'range_media_p50_reduction_percent':10,'both_repeats':True,'positive_regression_review_percent':5},'profile_order':[{'lane':lane,'kind':kind} for lane in [1,3] for kind in ['stat','record']],'claims':'Matched complete source-backed PPTX lifecycle; simulated range service, warm OS cache, no native application roundtrip or cold-disk/scaling claim. Process perf/RSS includes untimed fixture generation, gates and output hashing.','corpora':json.loads((ROOT.parent/'change-0448/protocol.json').read_text())['corpora'],'bound_files':{name:sha(name) for name in ['baseline-build.json','candidate-build.json','capture.py','verify-report.py','base-verify-report.py']}}
(ROOT/'protocol.json').write_text(json.dumps(v,indent=2)+'\n')
m={'uname':list(platform.uname()),'lscpu':subprocess.check_output(['lscpu'],text=True),'rustc':subprocess.check_output(['rustc','+1.98.1','-Vv'],text=True),'meminfo':Path('/proc/meminfo').read_text(),'cargo_config':(ROOT.parents[3]/'.cargo/config.toml').read_text() if (ROOT.parents[3]/'.cargo/config.toml').exists() else None}
(ROOT/'machine.json').write_text(json.dumps(m,indent=2)+'\n');print('FROZEN')
