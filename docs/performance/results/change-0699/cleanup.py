#!/usr/bin/env python3
"""Remove only the checked isolated worktree and this batch's three scratch trees."""
import hashlib,json,shutil,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
assert sys.argv[1:]==['--apply']
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
base=json.loads((P/'baseline.json').read_text())
assert all(sha(ROOT/n)==h for n,h in base['source_sha256'].items())
work=ROOT.parent/'litchi-0699-work'
assert all(sha(work/n)==h for n,h in base['source_sha256'].items())
assert subprocess.check_output(['git','diff','--name-only'],cwd=work,text=True)==''
assert subprocess.check_output(['git','diff','--cached','--name-only'],cwd=work,text=True)==''
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=work,text=True).strip()==base['baseline_head']
for name in subprocess.check_output(['git','ls-files','--others','--exclude-standard'],cwd=work,text=True).splitlines():
 assert name.startswith(str(P.relative_to(ROOT))+'/probe/'),name
for phase in ['baseline','candidate']:
 row=json.loads((P/f'build-{phase}.json').read_text());assert sha(Path(row['binary']))==row['binary_sha256']
lock=sha(ROOT/'Cargo.lock')
records=[]
for target in [work,ROOT.parent/'litchi-target-0699',ROOT.parent/'litchi-0699-bin',ROOT.parent/'litchi-0699-profile']:
 assert target.is_dir() and not target.is_symlink()
 if target==work:subprocess.run(['git','worktree','remove','--force',str(work)],cwd=ROOT,check=True)
 else:shutil.rmtree(target)
 records.append(dict(path=str(target),removed=not target.exists()))
assert sha(ROOT/'Cargo.lock')==lock
(P/'cleanup.json').write_text(json.dumps(dict(removed=records,workspace_cargo_lock_preserved_sha256=lock),indent=2)+'\n')
print('Removed exact four owned scratch trees; root source and lock preserved.')
