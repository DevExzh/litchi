"""Seal this evidence batch with a private index, preserving shared conflicts."""
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PARENT = '9a108730aa5533678b376ae097cd526c26e92b93'
BRANCH = 'refs/heads/perf/0912-pptx-native-sha-dispatch'
REPORT = ROOT / 'docs/performance/0912-pptx-native-sha-dispatch.md'
INDEX_NAMES = ('BASELINE.md', 'HOTSPOTS.md', 'REPORT.md', 'GOAL_AUDIT.md')


def git(*args, env=None, data=None):
    return subprocess.check_output(['git', *args], cwd=ROOT, env=env, input=data)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def desc(path):
    return {'path': str(path.relative_to(ROOT)), 'bytes': path.stat().st_size, 'sha256': sha(path)}


def main():
    assert not (P / 'closure.json').exists()
    assert subprocess.run(['git', 'show-ref', '--verify', '--quiet', BRANCH], cwd=ROOT).returncode == 1
    for name in ('dispatch-analysis.json', 'replay-before-cleanup.json', 'replay-after-cleanup.json', 'cleanup.json', 'result-review.json'):
        assert json.loads((P / name).read_text())['status'] == 'pass', name
    spec = importlib.util.spec_from_file_location('replay_0912_seal', P / 'replay.py')
    replay = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(replay)
    review = json.loads((P / 'result-review.json').read_text())
    replay.check_tree(review)
    assert review['final_report']['sha256'] == sha(REPORT)
    assert not list(P.rglob('__pycache__'))
    paths = [REPORT] + [ROOT / 'docs/performance' / name for name in INDEX_NAMES]
    paths += sorted(f for f in P.rglob('*') if f.is_file() and f.suffix != '.pptx' and f.name != 'perf.data')
    assert all(not f.is_symlink() for f in paths)
    index = Path(git('rev-parse', '--git-path', 'index').decode().strip())
    if not index.is_absolute():
        index = ROOT / index
    shared_index_hash = sha(index)
    shared_head = git('rev-parse', 'HEAD').decode().strip()
    cherry_pick = git('rev-parse', 'CHERRY_PICK_HEAD').decode().strip()
    closure = {'status': 'pass', 'parent_commit': PARENT, 'branch': BRANCH, 'shared_head': shared_head, 'shared_cherry_pick': cherry_pick, 'shared_index_sha256': shared_index_hash, 'integration': 'pending unrelated iWork cherry-pick resolution', 'production_changes': False, 'reused_binary': json.loads((P / 'plan.json').read_text())['binary'], 'perf_children': 1, 'perf_admission': 'rejected', 'lost_perf_records': 304, 'gdb_witnesses': 2, 'native_hardware_sha_reachability': 'observed inside measured owner for both routes', 'native_cost_or_speedup_claim': False, 'replay_before_and_after_cleanup': 'pass', 'worktrees_created': 0, 'capture_working_directories_removed': 2, 'cleanup_logical_bytes': 130, 'local_archives_and_perf': [desc(f) for f in sorted(P.rglob('*')) if f.is_file() and (f.suffix == '.pptx' or f.name == 'perf.data')], 'files': [desc(f) for f in sorted(paths)]}
    with (P / 'closure.json').open('x') as stream:
        json.dump(closure, stream, indent=2, sort_keys=True)
        stream.write('\n')
    paths.append(P / 'closure.json')
    names = sorted(str(f.relative_to(ROOT)) for f in paths)
    with tempfile.TemporaryDirectory(prefix='litchi-0912-commit-index-') as directory:
        directory = Path(directory)
        env = dict(os.environ, GIT_INDEX_FILE=str(directory / 'index'))
        git('read-tree', PARENT, env=env)
        spec = directory / 'paths'
        spec.write_bytes(b'\0'.join(name.encode() for name in names) + b'\0')
        git('add', '-f', '--pathspec-from-file=' + str(spec), '--pathspec-file-nul', env=env)
        changed = sorted(git('diff', '--cached', '--name-only', '-z', PARENT, env=env).decode().strip('\0').split('\0'))
        assert changed == names, {'unexpected': sorted(set(changed) ^ set(names))}
        authored = [name for name in names if not name.endswith(('.txt', '.patch')) and '/callgrind.out' not in name]
        # Preserve hash-bound evidence, including existing final blank lines.
        git('-c', 'core.whitespace=trailing-space,space-before-tab,-blank-at-eof',
            'diff', '--cached', '--check', PARENT, '--', *authored, env=env)
        tree = git('write-tree', env=env).decode().strip()
        commit = git('commit-tree', tree, '-p', PARENT, env=env, data=b'perf: verify native PPTX snapshot SHA dispatch\n\nRetain the rejected native perf profile and two owner-gated debugger\nwitnesses on the existing binary. Document the direct owned-source\nspan-copy candidate, offline replay and temporary-root cleanup.\n').decode().strip()
        assert git('show', commit + ':' + str((P / 'closure.json').relative_to(ROOT))) == (P / 'closure.json').read_bytes()
        for row in closure['files']:
            data = git('show', commit + ':' + row['path'])
            assert len(data) == row['bytes'] and hashlib.sha256(data).hexdigest() == row['sha256'], row['path']
        assert sha(index) == shared_index_hash and git('rev-parse', 'HEAD').decode().strip() == shared_head
        assert git('rev-parse', 'CHERRY_PICK_HEAD').decode().strip() == cherry_pick
        git('update-ref', BRANCH, commit, '0' * 40)
    assert sha(index) == shared_index_hash
    print(json.dumps({'status': 'pass', 'commit': commit, 'branch': BRANCH, 'files': len(names), 'shared_index_unchanged': True, 'temporary_index_removed': True}, indent=2))


if __name__ == '__main__':
    main()
