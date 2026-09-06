#!/usr/bin/env python3
"""Freeze actual pilot-passing matched simulation parameters before retained runs."""
import datetime,hashlib,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
for i in range(4):assert json.loads((ROOT/f'checks/pilot-{i}-control.json').read_text())['status']=='pass'
assert json.loads((ROOT/'oracle-probes.json').read_text())['status']=='pass'
assert not (ROOT/'protocol.json').exists() and not (ROOT/'runs').exists()
lanes=[{'corpus':corpus,'paced':paced} for paced in [False,True] for corpus in ['plain','media-rich']]
order=[dict(lane,repeat='R1') for lane in lanes]+[dict(lane,repeat='R2') for lane in reversed(lanes)]
sha=lambda name:hashlib.sha256((ROOT/name).read_bytes()).hexdigest()
value={'change':447,'frozen_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'status':'frozen','cpu':2,'workers':1,'samples':30,'warmups':3,'order':order,'max_range':65536,'fixed_delay_us':200,'paced_transfer_bytes_per_second':26214400,'reports':8,'retained_samples':240,'profile_order':[{'lane':lane,'kind':kind} for lane in [1,3] for kind in ['stat','record']],'profile_scope':'whole fresh process including corpus/gates/setup/warmups/diagnostics/reporting; cycles omit blocked sleep, no timed-only CPU attribution','repeat_review_percent':5,'claims':'single-build range-simulation baseline; no production speedup, actual network rate, cold I/O, shared-link concurrency or allocator attribution','corpora':{corpus:{k:json.loads((ROOT/f'pilots/{i}/control/report.json').read_text())[k] for k in ['source_archive_sha256','source_archive_bytes','destination_archive_sha256','destination_archive_bytes','expected_output_sha256','expected_output_bytes']} for corpus,i in [('plain',0),('media-rich',2)]},'bound_files':{name:sha(name) for name in ['build.json','capture.py','verify-report.py','base-verify-report.py','data-path.md']}}
(ROOT/'protocol.json').write_text(json.dumps(value,indent=2)+'\n');print('FROZEN')
