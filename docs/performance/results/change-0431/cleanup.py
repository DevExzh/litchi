#!/usr/bin/env python3
"""Remove only the named task scratch trees after portable validation."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
GOAL_SHA = 'bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1'

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    receipt = ROOT / 'cleanup.json'
    assert not receipt.exists()
    goal = REPO / 'docs/GOAL.md'
    assert sha(goal) == GOAL_SHA
    bundles = [ROOT, ROOT.with_name('change-0431-first-attempt')]
    for bundle in bundles:
        r = json.loads((bundle / 'checks/precleanup-portable.json').read_text())
        assert r['status'] == 'pass' and r['exit_code'] == 0
    binaries = Path('/tmp/litchi-goal-0431-binaries')
    for bundle, name, role in [(ROOT, 'before', 'before'), (bundles[1], 'after', 'after'), (ROOT, 'after-v2', 'after')]:
        b = json.loads((bundle / ('build-' + role + '.json')).read_text())
        assert sha(binaries / name) == b['binary_sha256']
    roots = [binaries, Path('/tmp/litchi-goal-0431-drafts')]
    inventory = []
    for tree in roots:
        assert tree.is_dir() and not tree.is_symlink()
        for path in sorted(tree.rglob('*')):
            assert not path.is_symlink()
            if path.is_file():
                inventory.append({'path': str(path), 'bytes': path.stat().st_size, 'sha256': sha(path)})
    adr_before = subprocess.check_output(['git', 'rev-parse', 'HEAD:docs/adr'], cwd=REPO, text=True).strip()
    assert adr_before == 'c950b6c8be822561b498d7bbe87c460873dcbf49'
    for tree in roots:
        shutil.rmtree(tree)
    assert all(not tree.exists() for tree in roots)
    assert sha(goal) == GOAL_SHA
    assert (REPO / 'target').is_dir() and (REPO / 'tools/perf-baseline/target').is_dir()
    receipt.write_text(json.dumps({'status': 'pass', 'removed_roots': [str(p) for p in roots], 'files': inventory, 'goal_sha256': GOAL_SHA, 'adr_tree': adr_before, 'preserved': ['target', 'tools/perf-baseline/target'], 'scope': 'Only task-owned binaries and drafts removed; fuzz scratch was separately cleaned with retained receipts.'}, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'removed_files': len(inventory)}))

if __name__ == '__main__':
    main()
