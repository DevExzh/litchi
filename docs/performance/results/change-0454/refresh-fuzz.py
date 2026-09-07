#!/usr/bin/env python3
"""Refresh the unbuilt fuzz target while retaining its initial preparation proof."""
import datetime
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TASK = Path('/tmp/litchi-goal-0454-opc-fuzz')


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def main():
    proof_path = ROOT / 'checks/fuzz-prepared.json'
    raw_proof = proof_path.read_bytes()
    proof = json.loads(raw_proof)
    if proof['task'] != str(TASK) or TASK.is_symlink() or (TASK / 'target').exists():
        raise ValueError('expected the exclusive, not-yet-built fuzz preparation')
    for item in proof['inputs']:
        path = TASK / item['path']
        raw = path.read_bytes()
        if path.is_symlink() or len(raw) != item['bytes'] or sha(raw) != item['sha256']:
            raise ValueError('initial fuzz preparation changed: ' + item['path'])
    target = TASK / 'fuzz_targets/parse_opc.rs'
    before = target.read_bytes()
    after = (REPO / 'crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs').read_bytes()
    if before == after:
        print('fuzz target already matches current source')
        return
    with (ROOT / 'checks/fuzz-prepared-initial.json').open('xb') as output:
        output.write(raw_proof)
    with (ROOT / 'candidate/fuzz-initial-parse_opc.rs.txt').open('xb') as output:
        output.write(before)
    target.write_bytes(after)
    for item in proof['inputs']:
        if item['path'] == 'fuzz_targets/parse_opc.rs':
            item.update(bytes=len(after), sha256=sha(after))
    proof['refresh'] = {
        'initial_proof_sha256': sha(raw_proof),
        'driver_sha256': sha(Path(__file__).read_bytes()),
        'utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
    }
    proof_path.write_text(json.dumps(proof, indent=2, sort_keys=True) + '\n')
    print('refreshed fuzz target; original preparation and target retained')


if __name__ == '__main__':
    main()
