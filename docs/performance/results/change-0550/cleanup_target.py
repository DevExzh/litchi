"""Verify all measured binary bytes and remove only the owned target."""
from pathlib import Path
import datetime
import hashlib
import json
import os
import shutil

B = Path(__file__).resolve().parent


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


plan = json.loads((B / 'plan.json').read_text())
target = Path(plan['owned_paths'][0])
assert plan['owned_paths'] == ['/home/zhuhe/litchi-goal-0550-target']
assert target.is_dir() and not target.is_symlink()
assert not (B / 'cleanup.json').exists()
binary_hashes = {}
for stage in ['baseline']:
    for kind in ['normal', 'alloc']:
        folder = B / stage
        identity = json.loads((folder / ('binary-' + kind + '.json')).read_text())
        binary = Path(identity['path'])
        assert binary.is_relative_to(target) and sha(binary) == identity['sha256']
        binary_hashes[kind] = identity['sha256']
references = []
for proc in Path('/proc').iterdir():
    if not proc.name.isdigit():
        continue
    values = []
    try:
        values.append((proc / 'cmdline').read_bytes().replace(b'\0', b' ').decode(errors='replace'))
    except OSError:
        pass
    for link in [proc / 'cwd', proc / 'exe']:
        try:
            values.append(os.readlink(link))
        except OSError:
            pass
    try:
        for link in (proc / 'fd').iterdir():
            try:
                values.append(os.readlink(link))
            except OSError:
                pass
    except OSError:
        pass
    matched = [value for value in values if str(target) in value]
    if matched:
        references.append({'pid': int(proc.name), 'references': matched})
assert not references, references
assert not list(B.rglob('__pycache__'))
shutil.rmtree(target)
assert not target.exists()
inputs = ['plan.json', 'run.py', 'capture.py', 'frozen-inputs.json']
receipt = {'target': str(target), 'removed': plan['owned_paths'],
           'owned_paths_absent': True, 'accessible_process_references': [],
           'observed_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
           'plan_sha256': sha(B / 'plan.json'),
           'binary_sha256_by_kind': binary_hashes,
           'python_cache_absent': True,
           'input_sha256': {name: sha(B / name) for name in inputs},
           'scope': 'All capture, quality and analysis processes terminal; strict precleanup passes. Two retained binary hashes checked before removal; accessible process command/cwd/exe/fd scan found no owned-target references.'}
(B / 'cleanup.json').write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n')
print('Verified two measured binaries and removed owned target.')
