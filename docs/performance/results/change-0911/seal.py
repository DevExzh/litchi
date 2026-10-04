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
PARENT = '52a27130589915884be5fd10bb1ade2f835ed32e'
BRANCH = 'refs/heads/perf/0911-pptx-owning-snapshot-profile'
REPORT = ROOT / 'docs/performance/0911-pptx-owning-snapshot-instruction-profile.md'
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
    for name in ('analysis.json', 'replay-before-cleanup.json', 'replay-after-cleanup.json', 'cleanup.json', 'result-review.json'):
        assert json.loads((P / name).read_text())['status'] == 'pass', name
    spec = importlib.util.spec_from_file_location('replay_0911_seal', P / 'replay.py')
    replay = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(replay)
    review = json.loads((P / 'result-review.json').read_text())
    replay.check_tree(review)
    review_paths = {'final_report': REPORT, 'canonical_analysis': P / 'analysis.json', 'analyzer': P / 'analyze.py', 'mutation_checker': P / 'analyzer-checks.py', 'protocol_review_v3': P / 'protocol-review-v3.md', 'candidate_assessment': P / 'candidate-assessment.md', 'plan_v3': P / 'plan-v3.json', 'replay_before_cleanup': P / 'replay-before-cleanup.json', 'replay_after_cleanup': P / 'replay-after-cleanup.json', 'cleanup': P / 'cleanup.json'}
    assert set(review['hash_bindings']) == set(review_paths)
    for name, path in review_paths.items():
        assert sha(path) == review['hash_bindings'][name], name
    assert sha(P / 'analyzer-checks.json') == review['fail_closed_mutations']['receipt']['sha256']
    assert not list(P.rglob('__pycache__'))
    paths = [REPORT] + [ROOT / 'docs/performance' / name for name in INDEX_NAMES]
    paths += sorted(f for f in P.rglob('*') if f.is_file() and f.suffix != '.pptx')
    assert all(not f.is_symlink() for f in paths)
    index = Path(git('rev-parse', '--git-path', 'index').decode().strip())
    if not index.is_absolute():
        index = ROOT / index
    shared_index_hash = sha(index)
    shared_head = git('rev-parse', 'HEAD').decode().strip()
    cherry_pick = git('rev-parse', 'CHERRY_PICK_HEAD').decode().strip()
    closure = {'status': 'pass', 'parent_commit': PARENT, 'branch': BRANCH, 'shared_head': shared_head, 'shared_cherry_pick': cherry_pick, 'shared_index_sha256': shared_index_hash, 'integration': 'pending unrelated iWork cherry-pick resolution', 'production_changes': False, 'reused_binary': json.loads((P / 'plan-v3.json').read_text())['binary'], 'admitted_children': 4, 'measured_contexts': 4, 'excluded_setup_contexts': 4, 'failed_attempts_retained': 2, 'native_latency_or_speedup_claim': False, 'replay_before_and_after_cleanup': 'pass', 'worktrees_created': 0, 'capture_working_directories_removed': 3, 'cleanup_logical_bytes': 195, 'local_archives': [desc(f) for f in sorted(P.rglob('*.pptx'))], 'files': [desc(f) for f in sorted(paths)]}
    with (P / 'closure.json').open('x') as stream:
        json.dump(closure, stream, indent=2, sort_keys=True)
        stream.write('\n')
    paths.append(P / 'closure.json')
    names = sorted(str(f.relative_to(ROOT)) for f in paths)
    with tempfile.TemporaryDirectory(prefix='litchi-0911-commit-index-') as directory:
        directory = Path(directory)
        env = dict(os.environ, GIT_INDEX_FILE=str(directory / 'index'))
        git('read-tree', PARENT, env=env)
        spec = directory / 'paths'
        spec.write_bytes(b'\0'.join(name.encode() for name in names) + b'\0')
        git('add', '-f', '--pathspec-from-file=' + str(spec), '--pathspec-file-nul', env=env)
        changed = sorted(git('diff', '--cached', '--name-only', '-z', PARENT, env=env).decode().strip('\0').split('\0'))
        assert changed == names, {'unexpected': sorted(set(changed) ^ set(names))}
        authored = [name for name in names if not name.endswith(('.txt', '.patch')) and '/callgrind.out' not in name]
        git('diff', '--cached', '--check', PARENT, '--', *authored, env=env)
        tree = git('write-tree', env=env).decode().strip()
        commit = git('commit-tree', tree, '-p', PARENT, env=env, data=b'perf: qualify PPTX cold snapshot instruction attribution\n\nRetain two failed scope attempts and four admitted caller-classified\nprofiles on the existing owning/materialized benchmark binary. Preserve\nsoftware SHA dispatch limits, offline replay and temporary-root cleanup.\n').decode().strip()
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
