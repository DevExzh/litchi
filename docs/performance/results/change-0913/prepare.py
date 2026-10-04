"""Reconstruct the verified 0910 source before applying the bounded candidate."""
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tarfile

ROOT = Path(__file__).resolve().parents[4]
P = Path(__file__).resolve().parent
OLD = ROOT / 'docs/performance/results/change-0910'
SOURCE = Path('/tmp/litchi-0913-src')
TARGET = ROOT / 'target/litchi-0913-build'

def sha(data):
    return hashlib.sha256(data).hexdigest()

def main():
    assert not SOURCE.exists() and not TARGET.exists()
    base = json.loads((OLD / 'source-base.json').read_text())
    frozen = json.loads((OLD / 'source-v1.json').read_text())
    archive = subprocess.check_output(['git', 'archive', base['commit'], '--', *base['archive_paths']], cwd=ROOT)
    assert sha(archive) == base['archive_sha256']
    SOURCE.mkdir()
    with tarfile.open(fileobj=io.BytesIO(archive)) as stream:
        stream.extractall(SOURCE, filter='data')
    for name, row in base['lock_overlays'].items():
        data = (ROOT / row['artifact']).read_bytes()
        assert sha(data) == row['sha256']
        (SOURCE / name).write_bytes(data)
    (SOURCE / '.owned-0910').write_text(base['commit'] + '\n')
    patch = ROOT / frozen['overlay_patch']['path']
    assert sha(patch.read_bytes()) == frozen['overlay_patch']['sha256']
    subprocess.run(['git', 'apply', str(patch)], cwd=SOURCE, check=True)
    actual = [{'path': str(f.relative_to(SOURCE)), 'bytes': f.stat().st_size, 'sha256': sha(f.read_bytes())} for f in sorted(SOURCE.rglob('*')) if f.is_file()]
    assert actual == frozen['files']
    TARGET.mkdir()
    for d in (SOURCE, TARGET):
        (d / '.owned-0913').write_text('owned-direct-read-0913\n')
    receipt = {'status': 'pass', 'reconstructed_files': len(actual), 'baseline_manifest_sha256': sha((OLD/'source-v1.json').read_bytes()), 'source_root': str(SOURCE), 'target_root': str(TARGET), 'worktrees_created': 0}
    with (P/'prepare.json').open('x') as out:
        json.dump(receipt,out,indent=2); out.write('\n')
    print(json.dumps(receipt))

if __name__ == '__main__':
    main()
