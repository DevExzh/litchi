#!/usr/bin/env python3
"""Bind the pilot-validated protocol to its final oracle before formal capture."""
import datetime,hashlib,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
p=ROOT/'protocol.json';v=json.loads(p.read_text());assert v['status']=='draft'
assert json.loads((ROOT/'checks/evidence-preflight.json').read_text())['status']=='pass'
assert json.loads((ROOT/'checks/independent-oracle-probes.json').read_text())['status']=='pass'
v['status']='frozen';v['frozen_utc']=datetime.datetime.now(datetime.timezone.utc).isoformat()
v['oracle']['verifier_sha256']=hashlib.sha256((ROOT/'oracle/verify-report.py').read_bytes()).hexdigest()
v['oracle']['protocol_path']='oracle/protocol.json';v['oracle']['protocol_sha256']=hashlib.sha256((ROOT/'oracle/protocol.json').read_bytes()).hexdigest()
v['analysis']={'p50':'arithmetic midpoint of the two central observations','p95_p99':'nearest rank','bootstrap_95':'2000 deterministic nonparametric resamples per report','repeat_review_trigger_percent':5,'scope':'descriptive single-baseline observations; allocator elapsed time is not a latency claim'}
p.write_text(json.dumps(v,indent=2)+'\n');print(json.dumps({'status':'frozen','protocol_sha256':hashlib.sha256(p.read_bytes()).hexdigest()}))
