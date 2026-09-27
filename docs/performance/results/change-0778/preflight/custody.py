"""Source and artifact custody for the 0778 ordinary-save experiment."""
from pathlib import Path
import hashlib
import json
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def sha(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def read(path):
    return json.loads(Path(path).read_text())


def write(path, value):
    Path(path).write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')


def census():
    paths = ['crates', 'tools/perf-baseline', 'Cargo.toml', 'clippy.toml', '.cargo/config.toml', 'rust-toolchain.toml']
    names = subprocess.check_output(['git', 'ls-files', '-z', '--', *paths], cwd=ROOT).decode().split('\0')
    files = {name: sha(ROOT / name) for name in names if name}
    files['Cargo.lock'] = sha(ROOT / 'Cargo.lock')
    return {'revision': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip(), 'files': files}


def artifact(path):
    path = Path(path)
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': sha(path)}
