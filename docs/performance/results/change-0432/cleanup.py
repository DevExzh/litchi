#!/usr/bin/env python3
"""Remove this batch's copied binaries after a passing portable replay."""
import hashlib
import json
from pathlib import Path
import shutil
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[3]
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def main():
    assert not (ROOT/'cleanup.json').exists()
    replay=json.loads((ROOT/'checks/precleanup-portable.json').read_text())
    assert replay['status']=='pass' and replay['exit_code']==0
    goal=REPO/'docs/GOAL.md'; expected='bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1'
    assert sha(goal)==expected
    tree=Path('/tmp/litchi-goal-0432-binaries'); assert tree.is_dir() and not tree.is_symlink()
    build=json.loads((ROOT/'build.json').read_text()); inventory=[]
    for mode,binary in build['binaries'].items():
        path=Path(binary['path']); assert path.parent==tree and path.name==mode and not path.is_symlink()
        assert sha(path)==binary['sha256'] and path.stat().st_size==binary['bytes']
        inventory.append({'path':str(path),'sha256':sha(path),'bytes':path.stat().st_size})
    assert {p.name for p in tree.iterdir()}==set(build['binaries'])
    shutil.rmtree(tree); assert not tree.exists() and sha(goal)==expected
    assert (REPO/'target').is_dir() and (REPO/'tools/perf-baseline/target').is_dir()
    (ROOT/'cleanup.json').write_text(json.dumps({'status':'pass','removed':inventory,'goal_sha256':expected,'preserved':['target','tools/perf-baseline/target'],'scope':'Only copied task binaries removed; portable replay does not require them.'},indent=2)+'\n')
    print(json.dumps({'status':'pass','removed_files':len(inventory)}))
if __name__=='__main__': main()
