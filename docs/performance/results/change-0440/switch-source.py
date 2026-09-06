#!/usr/bin/env python3
"""Restore the exact owned parser file between terminal ABBA CPU jobs."""
import argparse
import datetime
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
NAME = 'crates/litchi-odp/src/codec/parser/codec/xml/validation.rs'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--role', choices=('before', 'after'), required=True)
    parser.add_argument('--tag')
    args = parser.parse_args()
    previous = 'after' if args.role == 'before' else 'before'
    builds = {role: json.loads((ROOT / role / 'build.json').read_text()) for role in ('before', 'after')}
    sources = {role: json.loads((ROOT / builds[role]['source_manifest']['path']).read_text()) for role in builds}
    target = REPO / NAME
    digest = lambda raw: hashlib.sha256(raw).hexdigest()
    old = target.read_bytes()
    replacement = (ROOT / 'candidate' / (args.role + '-validation.rs.txt')).read_bytes()
    assert digest(old) == sources[previous][NAME]
    assert digest(replacement) == sources[args.role][NAME]
    tag = args.tag or args.role
    assert tag.replace('-', '').isalnum()
    proof = ROOT / 'switches' / (tag + '.json')
    assert not proof.exists()
    target.write_bytes(replacement)
    spec = importlib.util.spec_from_file_location('custody440', ROOT / 'check.py')
    custody = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(custody)
    after = custody.sources()
    assert after == builds[args.role]['source_manifest']
    proof.parent.mkdir(exist_ok=True)
    proof.write_text(json.dumps({'status':'pass', 'role':args.role, 'file':NAME, 'previous_sha256':digest(old), 'restored_sha256':digest(replacement), 'source_manifest':after, 'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(), 'scope':'root-owned source switch between terminal CPU jobs for matched ABBA capture'}, indent=2)+'\n')
    print('Restored exact ' + args.role + ' source')


if __name__ == '__main__':
    main()
