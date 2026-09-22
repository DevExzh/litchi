#!/usr/bin/env python3
"""Bind successful current-source qualification before prospective freeze."""
import hashlib,json
from pathlib import Path
P=Path(__file__).resolve().parent
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def latest(prefix):return max((p for p in P.glob(prefix+'-*') if p.is_dir()),key=lambda p:int(p.name.split('-')[-1]))
assert not (P/'freeze.json').exists()
b=read(P/'builds.json');probe={str(p.relative_to(P)):sha(p) for p in (P/'probe').rglob('*') if p.is_file()};assert probe==b['probe_sha256']
q=latest('quality');quality=read(q/'manifest.json');assert quality['probe_sha256']==probe and len(quality['runs'])==4 and all(r['exit_code']==0 and sha(q/r['log'])==r['sha256'] for r in quality['runs'])
s=latest('qualification');smoke=read(s/'manifest.json');assert smoke['builds_sha256']==sha(P/'builds.json') and smoke['script_sha256']==sha(P/'qualify.py') and len(smoke['runs'])==18
contract={}
for r in smoke['runs']:
 assert r['exit_code']==0 and sha(s/r['output'])==r['sha256'] and sha(s/r['stderr'])==r['stderr_sha256']
 x=read(s/r['output']);assert x['changed_length_proof']['logical_stream_length_change_proven']
 for oracle in [x['expected_oracle']]+[item['oracle'] for item in x['samples']]:assert all(v for v in oracle.values() if isinstance(v,bool)) and not oracle['failure_reasons']
 for control in x['oracle_controls']:assert control['rejected'] and control['status']=='rejected' and control['failure_reasons']
 assert x['changed_length_proof']['format_specific_semantic_length_proven']
 value=dict(directory_metadata_fields=x['directory_metadata_fields'],allocation_ownership_contract=x['allocation_ownership_contract'],semantic_witness=x['expected_oracle']['semantic_witness'],control_names=[c['name'] for c in x['oracle_controls']],identity={k:x[k] for k in ['source_sha256','expected_output_sha256','replacements_sha256','source_inventory','expected_output_inventory','replacements','changed_length_proof']})
 assert value==contract.setdefault(x['case'],value)
(P/'oracle-contract.json').write_text(json.dumps(contract,indent=2)+'\n')
receipts={str(p.relative_to(P)):sha(p) for directory in [q,s] for p in directory.iterdir() if p.is_file()}
(P/'qualification.json').write_text(json.dumps(dict(builds_sha256=sha(P/'builds.json'),files=receipts),indent=2)+'\n')
plan=read(P/'plan-draft.json');plan['status']='prospective fixed current-source baseline matrix';(P/'plan.json').write_text(json.dumps(plan,indent=2)+'\n');print('PASS current source qualified; prospective plan prepared')
