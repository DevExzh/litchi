#!/usr/bin/env python3
"""Clean only the rejected candidate's owned build roots after evidence replay."""
import datetime
import json
from pathlib import Path
import shlex
import shutil
import subprocess
import sys

import experiment

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
SOURCE = Path('/tmp/litchi-0887-source')
TARGET = ROOT / 'target/litchi-0887-build'


def read(path):
    return json.loads(path.read_text())


def main():
    if (P / 'cleanup.json').exists():
        raise ValueError('cleanup already recorded; do not restart')
    for name in ('build.json', 'verification.json', 'capture-review.json'):
        if read(P / name)['status'] != 'pass':
            raise ValueError('prerequisite did not pass: ' + name)
    if read(P / 'decision.json')['decision'] != 'reject':
        raise ValueError('this cleanup is for the rejected 0887 candidate')
    subprocess.run([sys.executable, '-B', str(P / 'freeze.py'), '--verify', '--retained-only'], check=True, stdout=subprocess.PIPE, text=True)
    build = read(P / 'build.json')
    if Path(build['source_root']) != SOURCE or Path(build['target_root']) != TARGET:
        raise ValueError('unexpected owned build paths')
    for path in (SOURCE, TARGET):
        if path.is_symlink() or not path.is_dir() or (path / '.owner-0887').read_text() != 'MCE attribute candidate 0887\n':
            raise ValueError('invalid ownership marker: ' + str(path))
    manifest = read(P.parent / 'change-0885/source.json')
    candidate = read(P / 'candidate-source.json')
    for name in candidate['overlays']:
        if experiment._sha256(ROOT / name) != manifest['files'][name]['sha256']:
            raise ValueError('production was not restored to baseline: ' + name)
    dependencies = {}
    for role in ('native', 'allocator'):
        line = (P / (role + '.d')).read_text().splitlines()[0]
        paths = shlex.split(line.split(': ', 1)[1])
        for name in paths:
            relative = Path(name).resolve().relative_to(SOURCE).as_posix()
            if relative not in manifest['files']:
                raise ValueError('unmanifested source dependency: ' + name)
        dependencies[role] = len(paths)
    for path in (SOURCE, TARGET):
        shutil.rmtree(path)
        if path.exists():
            raise ValueError('cleanup failed: ' + str(path))
    experiment._write_json(P / 'cleanup.json', {
        'status': 'pass',
        'finished_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
        'removed_owned_scratch': [str(SOURCE), str(TARGET)],
        'source_dependency_counts_verified': dependencies,
        'registered_worktrees_created': 0,
        'production_restored_to_baseline': True,
        'retained_artifacts': [experiment.descriptor(P / 'candidate.patch'), *[experiment.descriptor(P / 'binaries' / role) for role in ('native', 'allocator')]],
        'scope': 'Only marked 0887 build/source roots; historical recovery and other teams worktrees preserved.',
    })
    print('Owned 0887 build roots removed; retained artifacts preserved.')


if __name__ == '__main__':
    main()
