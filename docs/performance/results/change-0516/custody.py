"""Replay each candidate patch on only its base blobs in temporary directories."""
import json
import subprocess
import tempfile
from pathlib import Path

from run import HERE, REPO, sha
from audit import read


def check():
    before = read(HERE / 'before/source-manifest.json')
    base = read(HERE / 'plan.json')['base_revision']
    results = []
    for record_path in sorted(HERE.glob('*/candidate-patch.json')):
        record = read(record_path)
        directory = record_path.parent
        patch = directory / 'candidate.patch'
        manifest_path = directory / 'source-manifest.json'
        assert record['base_revision'] == base
        assert sha(patch) == record['patch_sha256']
        assert sha(manifest_path) == record['source_manifest_sha256']
        manifest = read(manifest_path)
        changed = {name for name in before.keys() | manifest.keys() if before.get(name) != manifest.get(name)}
        assert changed == set(record['paths']) == set(record['candidate_file_sha256'])
        assert all(manifest[name] == value for name, value in record['candidate_file_sha256'].items())
        with tempfile.TemporaryDirectory(prefix='litchi-0516-patch-replay-') as temporary:
            target = Path(temporary)
            for name in sorted(changed):
                assert name.startswith('crates/litchi-xlsx/') and '..' not in Path(name).parts
                result = subprocess.run(['git', 'show', base + ':' + name], cwd=REPO, capture_output=True)
                if name in before:
                    assert result.returncode == 0
                    path = target / name
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_bytes(result.stdout)
                    assert sha(path) == before[name]
                else:
                    assert result.returncode != 0
            result = subprocess.run(['git', 'apply', '--whitespace=nowarn', str(patch)], cwd=target, capture_output=True)
            assert result.returncode == 0, result.stderr.decode()
            actual = {str(path.relative_to(target)): sha(path) for path in target.rglob('*') if path.is_file()}
            assert actual == record['candidate_file_sha256']
        assert not target.exists()
        results.append({'epoch': directory.name, 'changed_files': len(changed),
                        'patch_sha256': sha(patch), 'source_manifest_sha256': sha(manifest_path),
                        'base_blob_and_patch_replay': True, 'temporary_directory_removed': True})
    assert results
    return {'base_revision': base, 'epochs': results}


if __name__ == '__main__':
    print(json.dumps(check(), indent=2))
