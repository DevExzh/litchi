#!/usr/bin/env python3
"""Record the isolated fuzz-driver ownership correction after its failed build."""
import datetime
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TASK = Path('/tmp/litchi-goal-0454-opc-fuzz')
NAME = 'crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs'

def sha(raw):
    return hashlib.sha256(raw).hexdigest()

def write_new(path, value):
    with path.open('x') as stream:
        stream.write(json.dumps(value, indent=2) + '\n')

failed = json.loads((ROOT / 'checks/fuzz-build.json').read_text())
if failed['status'] != 'failed' or not failed['source_unchanged']:
    raise ValueError('expected completed failed fuzz build with stable sources')
spec = importlib.util.spec_from_file_location('custody', ROOT / 'check.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
before = json.loads((ROOT / 'candidate-build.json').read_text())['source_manifest']
after = module.sources()
old = json.loads((ROOT / before['path']).read_text())
new = json.loads((ROOT / after['path']).read_text())
changed = [name for name in old.keys() | new.keys() if old.get(name) != new.get(name)]
if changed != [NAME]:
    raise ValueError(f'expected only the standalone fuzz target to change: {changed}')
proof_path = ROOT / 'checks/fuzz-prepared.json'
raw_proof = proof_path.read_bytes()
proof = json.loads(raw_proof)
for item in proof['inputs']:
    path = TASK / item['path']
    raw = path.read_bytes()
    if path.is_symlink() or len(raw) != item['bytes'] or sha(raw) != item['sha256']:
        raise ValueError('prepared fuzz input changed: ' + item['path'])
target = TASK / 'fuzz_targets/parse_opc.rs'
if sha(target.read_bytes()) != old[NAME]:
    raise ValueError('temporary target differs from measured source manifest')
with (ROOT / 'checks/fuzz-prepared-before-amendment.json').open('xb') as stream:
    stream.write(raw_proof)
raw = (REPO / NAME).read_bytes()
snapshot = ROOT / 'candidate/fuzz-amended-parse_opc.rs.txt'
with snapshot.open('xb') as stream:
    stream.write(raw)
target.write_bytes(raw)
for item in proof['inputs']:
    if item['path'] == 'fuzz_targets/parse_opc.rs':
        item.update(bytes=len(raw), sha256=sha(raw))
proof['amendment'] = 'fuzz-source-amendment.json'
proof_path.write_text(json.dumps(proof, indent=2, sort_keys=True) + '\n')
write_new(ROOT / 'fuzz-source-amendment.json', {
    'status': 'pass', 'before': before, 'after': after,
    'allowed_changed_path': NAME, 'changed_paths': changed,
    'before_sha256': old[NAME], 'after_sha256': new[NAME],
    'snapshot': str(snapshot.relative_to(ROOT)),
    'reason': 'Accumulate the topology plan while borrowing parts; consume the package only after iteration. The standalone fuzz target is outside production binary dependencies.',
    'recorded_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
})
print('recorded isolated fuzz target amendment')
