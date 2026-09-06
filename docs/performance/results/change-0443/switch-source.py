#!/usr/bin/env python3
"""Switch the explicitly inventoried candidate files between terminal CPU jobs."""
import argparse
import datetime
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--role', choices=('before', 'after'), required=True)
    parser.add_argument('--tag')
    args = parser.parse_args()
    previous = 'after' if args.role == 'before' else 'before'
    names = json.loads((ROOT / 'source-files.json').read_text())
    assert len(names) == len(set(Path(name).name for name in names))
    builds = {role: json.loads((ROOT / role / 'build.json').read_text()) for role in ('before', 'after')}
    sources = {role: json.loads((ROOT / build['source_manifest']['path']).read_text()) for role, build in builds.items()}
    digest = lambda raw: hashlib.sha256(raw).hexdigest()
    replacements = {}
    for name in names:
        target = REPO / name
        assert name.startswith('crates/litchi-odp/src/') and name.endswith('.rs')
        assert target.resolve().is_relative_to(REPO) and not target.is_symlink()
        actual = digest(target.read_bytes()) if target.exists() else None
        assert actual == sources[previous].get(name), name
        candidate = ROOT / 'candidate' / (args.role + '-' + target.name + '.txt')
        if name in sources[args.role]:
            raw = candidate.read_bytes()
            assert digest(raw) == sources[args.role][name]
            replacements[name] = raw
        else:
            assert not candidate.exists()
            replacements[name] = None
    tag = args.tag or args.role
    assert tag.replace('-', '').isalnum()
    proof = ROOT / 'switches' / (tag + '.json')
    assert not proof.exists()
    for name, raw in replacements.items():
        if raw is None:
            (REPO / name).unlink()
        else:
            (REPO / name).write_bytes(raw)
    spec = importlib.util.spec_from_file_location('custody443', ROOT / 'check.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    current = custody.sources()
    assert current == builds[args.role]['source_manifest']
    proof.parent.mkdir(exist_ok=True)
    proof.write_text(json.dumps({'status':'pass','role':args.role,'files':names,'source_manifest':current,'utc':datetime.datetime.now(datetime.timezone.utc).isoformat()}, indent=2) + '\n')
    print('Restored exact ' + args.role + ' source')


if __name__ == '__main__':
    main()
