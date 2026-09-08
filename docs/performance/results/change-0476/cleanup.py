#!/usr/bin/env python3
"""Remove only the runtime trees and binary copies created for this experiment."""
from pathlib import Path
import shutil
import subprocess
from common import ROOT, REPO, TEMP, now, sha, write, check_arm


def main():
    record = dict(started_utc=now(), driver_sha256=sha(Path(__file__)),
        owned_roots=[str(TEMP), '/tmp/litchi-goal-0474'],
        retained_caches=[str(REPO / 'target'), str(REPO / 'tools/perf-baseline/target')])
    user_files = {
        'docs/GOAL.md': 'bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1',
        'docs/report/spec-gap-audit.md': 'f71288dddc0827d5f5a16bf3936985606ec1b341d6e5c06fced4e612c7972ddf',
    }
    for path, digest in user_files.items():
        assert sha(REPO / path) == digest, path
    builds = [check_arm(arm) for arm in ('control', 'candidate')]
    for build in builds:
        tree = Path(build['build_path'])
        assert tree in (TEMP / 'candidate', Path('/tmp/litchi-goal-0474/tree'))
        subprocess.run(['git', 'worktree', 'remove', str(tree)], cwd=REPO, check=True)
    for directory in (TEMP, Path('/tmp/litchi-goal-0474')):
        assert directory.is_dir() and not directory.is_symlink()
        shutil.rmtree(directory)
        assert not directory.exists()
    record.update(finished_utc=now(), all_owned_roots_absent=True,
        user_files_unchanged={p: sha(REPO / p) == h for p, h in user_files.items()},
        retained_caches_present={p: Path(p).is_dir() for p in record['retained_caches']})
    assert all(record['user_files_unchanged'].values())
    assert all(record['retained_caches_present'].values())
    write(ROOT / 'cleanup.json', record)
    print('owned runtime cleanup passed')


if __name__ == '__main__':
    main()
