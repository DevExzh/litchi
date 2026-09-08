#!/usr/bin/env python3
"""Verify current source/binaries, replay in a fresh copy, and clean owned binaries."""
import datetime
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import capture

ROOT = Path(__file__).resolve().parent


def save(name, value):
    (ROOT / name).write_text(json.dumps(value, indent=2) + '\n')


def seal():
    rows = []
    for path in sorted(ROOT.rglob('*')):
        assert not path.is_symlink(), path
        if path.is_file() and path != ROOT / 'SHA256SUMS':
            assert path.suffix != '.pyc' and '__pycache__' not in path.parts
            rows.append(f'{capture.sha(path)}  {path.relative_to(ROOT)}\n')
    (ROOT / 'SHA256SUMS').write_text(''.join(rows))


def verify(root):
    result = subprocess.run([sys.executable, '-B', str(root / 'verify.py')], cwd=root,
                            text=True, capture_output=True)
    assert result.returncode == 0, result.stdout + result.stderr
    receipt = json.loads(result.stdout)
    assert receipt['status'] == 'pass'
    return receipt


def main():
    binding = capture.source_check()
    profile = json.loads((ROOT / 'profile-binding.json').read_text())
    owned = {capture.TEMP / 'normal': binding['sha256'],
             capture.TEMP / 'profile': profile['binary_sha256']}
    assert set(capture.TEMP.iterdir()) == set(owned)
    entries = []
    for path, digest in owned.items():
        assert path.is_file() and not path.is_symlink() and capture.sha(path) == digest
        entries.append(dict(path=str(path), bytes=path.stat().st_size, sha256=digest))
    seal()
    before = verify(ROOT)
    save('precleanup.json', dict(status='pass', verifier_sha256=capture.sha(ROOT / 'verify.py'),
                                source_files=7032, binaries=entries, verification=before))
    seal()
    with tempfile.TemporaryDirectory(prefix='litchi-0466-portable-') as temp_name:
        parent = Path(temp_name)
        target = parent / ROOT.name
        shutil.copytree(ROOT, target)
        shutil.copytree(ROOT.parent / 'change-0465', parent / 'change-0465')
        result = verify(target)
        assert capture.sha(target / 'SHA256SUMS') == capture.sha(ROOT / 'SHA256SUMS')
    assert not Path(temp_name).exists()
    save('portable-verification.json', dict(status='pass', verification=result,
                                           verifier_sha256=capture.sha(ROOT / 'verify.py'),
                                           temporary_copy_removed=True,
                                           scope='fresh-copy replay with adjacent frozen 0465 provenance'))
    for path, digest in owned.items():
        assert capture.sha(path) == digest
        path.unlink()
    capture.TEMP.rmdir()
    save('cleanup.json', dict(status='pass', finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                              files_removed=entries, removed_bytes=sum(row['bytes'] for row in entries),
                              task_directory_absent=not capture.TEMP.exists()))
    seal()
    print(json.dumps(verify(ROOT), indent=2))


if __name__ == '__main__':
    main()
