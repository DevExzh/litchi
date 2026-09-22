#!/usr/bin/env python3
"""Bind final probe qualification and inherited DOC semantics before freeze."""
import hashlib,json
from pathlib import Path
P=Path(__file__).resolve().parent
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def latest(prefix):return max((p for p in P.glob(prefix+'-*') if p.is_dir()),key=lambda p:int(p.name.split('-')[-1]))
assert not (P/'freeze.json').exists();b=read(P/'builds.json');probe={str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()};assert probe==b['probe_sha256']
q=latest('quality');quality=read(q/'manifest.json');assert quality['probe_sha256']==probe and len(quality['runs'])==4 and all(r['exit_code']==0 and sha(q/r['log'])==r['sha256'] for r in quality['runs'])
s=latest('qualification');smoke=read(s/'manifest.json');assert smoke['builds_sha256']==sha(P/'builds.json') and smoke['script_sha256']==sha(P/'qualify.py') and len(smoke['runs'])==8
prior=read(P/'prior-oracle-contract.json');contract={}
parent=P.parent/'change-0728';ancestry=read(P/'ancestry.json');assert sha(parent/'artifact-manifest.json')==ancestry['sealed_packet_manifest_sha256'] and sha(parent/'oracle-contract.json')==ancestry['prior_oracle_contract_sha256']
assert prior=={k:v for k,v in read(parent/'oracle-contract.json').items() if k in ['docfloat','docnohf']}
assert b['source_sha256']==read(parent/'builds.json')['source_sha256']
for r in smoke['runs']:
 assert r['exit_code']==0 and sha(s/r['output'])==r['sha256'] and sha(s/r['stderr'])==r['stderr_sha256']
 x=read(s/r['output']);identity={k:x[k] for k in prior[x['case']]['identity']};assert identity==prior[x['case']]['identity']
 assert x['expected_oracle']['semantic_witness']==prior[x['case']]['semantic_witness']
 for oracle in [x['expected_oracle']]+[item['oracle'] for item in x['samples']]:assert all(v for v in oracle.values() if isinstance(v,bool)) and not oracle['failure_reasons']
 for control in x['oracle_controls']:assert control['rejected'] and control['status']=='rejected' and control['failure_reasons']
 value=dict(directory_metadata_fields=x['directory_metadata_fields'],allocation_ownership_contract=x['allocation_ownership_contract'],identity=identity,semantic_witness=x['expected_oracle']['semantic_witness'],control_names=[c['name'] for c in x['oracle_controls']],headers={k:x[k] for k in ['phase_contract','diagnostic_contract','observer_contract']})
 assert value==contract.setdefault(x['case'],value)
(P/'oracle-contract.json').write_text(json.dumps(contract,indent=2)+'\n')
receipts={str(p.relative_to(P)):sha(p) for directory in [q,s] for p in directory.iterdir() if p.is_file()}
(P/'qualification.json').write_text(json.dumps(dict(builds_sha256=sha(P/'builds.json'),files=receipts),indent=2)+'\n')
plan=read(P/'plan-draft.json');plan['status']='prospective fixed public DOC attribution matrix';(P/'plan.json').write_text(json.dumps(plan,indent=2)+'\n');print('PASS final source qualified; exact 0728 DOC oracle and prospective plan')
