#!/usr/bin/env python3
"""Remove only the three batch-0434 scratch directories after portable proof."""
import hashlib
import json
from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
PATHS = tuple(Path('/tmp') / name for name in (
    'litchi-goal-0434-binaries', 'litchi-goal-0434-ods',
    'litchi-goal-0434-ods-tests',
))

def main():
    receipt = json.loads((ROOT / 'checks/precleanup-portable-v2.json').read_text())
    assert receipt['status'] == 'pass' and receipt['exit_code'] == 0
    output = ROOT / 'checks/cleanup-inventory.json'
    assert not output.exists()
    keep = [REPO / 'target', REPO / 'tools/perf-baseline/target']
    before = {str(p): [p.stat().st_dev, p.stat().st_ino] for p in keep}
    goal = REPO / 'docs/GOAL.md'
    goal_hash = hashlib.sha256(goal.read_bytes()).hexdigest()
    assert goal_hash == 'bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1'
    rows = []
    for path in PATHS:
        assert path.is_dir() and not path.is_symlink()
        files = [p for p in path.rglob('*') if p.is_file() and not p.is_symlink()]
        rows.append({'path': str(path), 'regular_files': len(files),
                     'regular_file_bytes': sum(p.stat().st_size for p in files)})
    for path in PATHS:
        shutil.rmtree(path)
        assert not path.exists()
    assert before == {str(p): [p.stat().st_dev, p.stat().st_ino] for p in keep}
    assert goal_hash == hashlib.sha256(goal.read_bytes()).hexdigest()
    output.write_text(json.dumps({'status': 'pass', 'removed': rows,
        'preserved_target_directory_identity': before,
        'user_goal_sha256': goal_hash}, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'removed_directories': len(rows),
                      'regular_file_bytes': sum(r['regular_file_bytes'] for r in rows)}))

if __name__ == '__main__':
    main()
