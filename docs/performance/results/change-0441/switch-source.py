#!/usr/bin/env python3
"""Switch only the two owned source files between terminal measurement jobs."""
import argparse
import datetime
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
NAMES = ['crates/litchi-odp/src/authoring/' + name for name in ('edit.rs', 'mutable.rs')]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--role', choices=('before', 'after'), required=True)
    parser.add_argument('--tag')
    args = parser.parse_args()
    previous = 'after' if args.role == 'before' else 'before'
    builds = {role: json.loads((ROOT / role / 'build.json').read_text()) for role in ('before', 'after')}
    sources = {role: json.loads((ROOT / build['source_manifest']['path']).read_text()) for role, build in builds.items()}
    digest = lambda raw: hashlib.sha256(raw).hexdigest()
    replacements = {}
    for name in NAMES:
        raw = (ROOT / 'candidate' / (args.role + '-' + Path(name).name + '.txt')).read_bytes()
        assert digest((REPO / name).read_bytes()) == sources[previous][name]
        assert digest(raw) == sources[args.role][name]
        replacements[name] = raw
    tag = args.tag or args.role
    assert tag.replace('-', '').isalnum()
    proof = ROOT / 'switches' / (tag + '.json')
    assert not proof.exists()
    for name, raw in replacements.items():
        (REPO / name).write_bytes(raw)
    spec = importlib.util.spec_from_file_location('custody441', ROOT / 'check.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    current = custody.sources()
    assert current == builds[args.role]['source_manifest']
    proof.parent.mkdir(exist_ok=True)
    proof.write_text(json.dumps({'status': 'pass', 'role': args.role, 'files': NAMES, 'source_manifest': current, 'utc': datetime.datetime.now(datetime.timezone.utc).isoformat()}, indent=2) + '\n')
    print('Restored exact ' + args.role + ' source')


if __name__ == '__main__':
    main()
