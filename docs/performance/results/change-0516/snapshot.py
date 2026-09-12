"""Freeze a reviewed isolated checkout into a new, reproducible source epoch."""
import argparse
import json
import shutil
import subprocess

from run import HERE, REPO, SCRATCH, sha, sources, write


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('epoch')
    parser.add_argument('--from-epoch', required=True)
    args = parser.parse_args()
    assert args.epoch.replace('-', '').isalnum()
    assert args.from_epoch.replace('-', '').isalnum()
    assert sources() == json.loads((HERE / args.from_epoch / 'source-manifest.json').read_text())
    worktree = SCRATCH / 'worktree'
    tracked = subprocess.check_output(['git', 'diff', '--name-only', '--', 'crates/litchi-xlsx'], cwd=worktree).decode().splitlines()
    new = subprocess.check_output(['git', 'ls-files', '--others', '--exclude-standard', '--', 'crates/litchi-xlsx'], cwd=worktree).decode().splitlines()
    paths = sorted(set(tracked + new))
    assert paths and all(path.startswith('crates/litchi-xlsx/') and path.endswith('.rs') for path in paths)
    hashes = {path: sha(worktree / path) for path in paths}
    directory = HERE / args.epoch
    assert not directory.exists()
    for path in paths:
        shutil.copy2(worktree / path, REPO / path)
    assert {path: sha(worktree / path) for path in paths} == hashes
    assert {path: sha(REPO / path) for path in paths} == hashes
    directory.mkdir()
    write(directory / 'source-manifest.json', sources())
    patch = subprocess.check_output(['git', 'diff', '--binary', '--', 'crates/litchi-xlsx'], cwd=REPO)
    for path in new:
        result = subprocess.run(['git', 'diff', '--no-index', '--binary', '--', '/dev/null', path], cwd=REPO, capture_output=True)
        assert result.returncode == 1
        patch += result.stdout
    (directory / 'candidate.patch').write_bytes(patch)
    write(directory / 'candidate-patch.json', {
        'base_revision': json.loads((HERE / 'plan.json').read_text())['base_revision'],
        'paths': paths, 'candidate_file_sha256': hashes,
        'patch_sha256': sha(directory / 'candidate.patch'),
        'source_manifest_sha256': sha(directory / 'source-manifest.json'),
        'previous_epoch': args.from_epoch,
        'scope': 'Source snapshot, independently bound to every test or capture receipt in this epoch.'})
    print(args.epoch, len(paths), 'files frozen')


if __name__ == '__main__':
    main()
