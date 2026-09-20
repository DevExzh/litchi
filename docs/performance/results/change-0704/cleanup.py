#!/usr/bin/env python3
"""Remove only this batch's named scratch after frozen evidence has passed."""
import hashlib,json,shutil,sys,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
assert sys.argv[1:]==['--apply']
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
lock=ROOT/'Cargo.lock'
lock_hash=sha(lock) if lock.exists() else None
targets=[ROOT.parent/'litchi-target-0704',ROOT.parent/'litchi-0704-bin',ROOT.parent/'litchi-0704-profile',P/'marker-control.pptx']
WITNESS=P/'candidate-source-witness.json'
witness=json.loads(WITNESS.read_text())
assert {'final_candidate_patch','final_candidate_patch_sha256','final_candidate_source_sha256','final_candidate_patch_files'}.issubset(witness)
handoff=witness.get('historical_handoff')
if handoff is None:
 saved=P/'preflight/candidate-coder-handoff-witness.json'
 handoff=json.loads(saved.read_text()) if saved.exists() else witness
patch_ref=handoff.get('candidate_patch','candidate-production.patch')
PATCH=(P/patch_ref).resolve()
assert PATCH.is_file() and PATCH.is_relative_to(P.resolve())
baseline=json.loads((P/'baseline.json').read_text())['baseline_head']
assert baseline.startswith(witness['baseline'])
assert sha(PATCH)==handoff['candidate_patch_sha256']
candidate_worktree=Path(handoff['candidate_worktree'])
assert candidate_worktree.is_dir() and not candidate_worktree.is_symlink()
candidate_sources=handoff['candidate_source_sha256']
assert candidate_sources
assert set(candidate_sources)==set(handoff['candidate_patch_files'])
assert all(name.startswith('crates/litchi-pptx/') for name in candidate_sources)
for name,digest in candidate_sources.items():
 path=candidate_worktree/name
 assert path.is_file() and sha(path)==digest, f'candidate witness mismatch: {path}'
targets.append(candidate_worktree)
retention=json.loads((P/'retention-build.json').read_text())
retention_binary=Path(retention['binary'])
assert retention_binary.is_file() and sha(retention_binary)==retention['binary_sha256']
mechanism_receipt=None
for name in ('mechanism/build.json','mechanism/mechanism-build.json','mechanism/build-mechanism.json'):
 path=P/name
 if path.is_file():
  assert mechanism_receipt is None, 'multiple mechanism build receipts'
  mechanism_receipt=json.loads(path.read_text())
if (P/'mechanism').exists():
 assert mechanism_receipt is not None, 'mechanism directory has no build receipt'
mechanism_binary=None
if mechanism_receipt is not None:
 mechanism_binary=Path(mechanism_receipt['binary'])
 assert mechanism_binary.is_file() and sha(mechanism_binary)==mechanism_receipt['binary_sha256']
for phase in ['baseline','candidate']:
 for row in json.loads((P/f'builds-{phase}.json').read_text())+[json.loads((P/f'build-refusal-{phase}.json').read_text())]:
  assert sha(Path(row['binary']))==row['binary_sha256']
records=[]
binaries=[]
for target in targets:
 assert not target.is_symlink()
 record=dict(path=str(target),existed=target.exists(),kind='directory' if target.is_dir() else 'file')
 if target == candidate_worktree:
  subprocess.run(['git','worktree','remove','--force',str(target)],cwd=ROOT,check=True)
 elif target.is_dir():shutil.rmtree(target)
 elif target.exists():
  record['sha256']=sha(target);target.unlink()
 record['removed']=not target.exists();records.append(record)
assert (sha(lock) if lock.exists() else None)==lock_hash
for phase in ['baseline','candidate']:
 for row in json.loads((P/f'builds-{phase}.json').read_text())+[json.loads((P/f'build-refusal-{phase}.json').read_text())]:
  binaries.append(dict(phase=phase,label=row.get('label','refusal'),path=row['binary'],binary_sha256=row['binary_sha256']))
binaries.append(dict(phase='candidate',label='retention-observer',path=str(retention_binary),binary_sha256=retention['binary_sha256']))
if mechanism_receipt is not None:
 binaries.append(dict(phase='candidate',label='mechanism',path=str(mechanism_binary),binary_sha256=mechanism_receipt['binary_sha256']))
(P/'cleanup.json').write_text(json.dumps(dict(removed=records,binaries=binaries,historical_worktree_witness=dict(path=str(candidate_worktree),candidate_patch=patch_ref,candidate_patch_sha256=handoff['candidate_patch_sha256'],candidate_source_sha256=candidate_sources),workspace_cargo_lock_preserved_sha256=lock_hash,scope='Exact batch-owned scratch only; shared targets and unrelated work preserved.'),indent=2)+'\n')
print('removed exactly five owned scratch paths; workspace Cargo.lock preserved')
