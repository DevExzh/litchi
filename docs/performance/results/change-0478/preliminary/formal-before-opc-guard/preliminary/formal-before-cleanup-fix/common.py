import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path('/tmp/litchi-goal-0478')
ENV = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4',
           CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_DEBUG='1',
           RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes',
           DEBUGINFOD_URLS='', LC_ALL='C', PYTHONDONTWRITEBYTECODE='1')
ENV_KEYS = ('RUSTUP_TOOLCHAIN', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL',
            'CARGO_PROFILE_RELEASE_DEBUG', 'RUSTFLAGS', 'DEBUGINFOD_URLS', 'LC_ALL')


def sha(path):
    with Path(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def read(path):
    return json.loads(Path(path).read_text())


def write(path, value):
    with Path(path).open('x') as stream:
        json.dump(value, stream, indent=2, sort_keys=True); stream.write('\n')


def meta(path):
    path = Path(path)
    return dict(bytes=path.stat().st_size, sha256=sha(path))


def check_source(source):
    tree = Path(source['build_path'])
    assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=tree, text=True).strip() == source['revision']
    assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=tree).strip()
    manifest = ROOT / source['source_manifest']['path']
    assert sha(manifest) == source['source_manifest']['sha256']
    for path, digest in read(manifest).items():
        assert sha(tree / path) == digest, path
    for path, digest in source['fixtures'].items():
        assert sha(tree / path) == digest, path


def check_arm(arm):
    build = read(ROOT / f'{arm}-build.json')
    check_source(build)
    for binary in build['binaries'].values():
        assert meta(binary['path']) == {k: binary[k] for k in ('bytes', 'sha256')}
    return build
