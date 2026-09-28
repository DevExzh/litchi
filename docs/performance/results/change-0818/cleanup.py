"""Remove only the checked 0818 target and marked scratch after validation."""
import json
from pathlib import Path
import shutil
import subprocess
import sys
import time

import run as custody

P, ROOT, PLAN = custody.P, custody.ROOT, custody.PLAN
assert not (P / 'cleanup.json').exists()
assert (P / 'results-review.md').is_file()
subprocess.run([sys.executable, '-B', str(P / 'validate.py')], cwd=ROOT, check=True)
custody.guard()
source_before = custody.source()
outcome = json.loads((P / 'outcome.json').read_text())
binary = outcome['binary']
assert custody.sha(Path(binary['path'])) == binary['sha256']
assert Path(binary['path']).stat().st_size == binary['bytes']
target, scratch = Path(PLAN['target']), Path(PLAN['scratch'])
assert target == ROOT.parent / 'litchi-target-0818'
assert scratch == ROOT.parent / 'litchi-fs-0818'
assert (scratch / '.litchi-performance-0818-owned').read_text() == 'litchi-performance-0818-owned-scratch-v1\n'
started = time.time()
removed = []
for path in (target, scratch):
    assert path.is_dir() and not path.is_symlink()
    files = [item for item in path.rglob('*') if item.is_file()]
    removed.append({'path': str(path), 'files': len(files), 'logical_bytes': sum(item.stat().st_size for item in files)})
    shutil.rmtree(path)
    assert not path.exists()
custody.guard()
assert custody.source() == source_before
custody.write(P / 'cleanup.json', {'schema': 'litchi.performance.0818.cleanup.v1', 'binary': binary, 'binary_verified_before_removal': True, 'removed': removed, 'started': started, 'ended': time.time()})
print('0818 owned target and scratch removed; source and unrelated files unchanged')
