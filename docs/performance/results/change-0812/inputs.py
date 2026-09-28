"""Read-only custody of the sealed diagnostic used for instruction localization."""
import hashlib
import json
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P.parent / 'change-0811'

def read(path):
    return json.loads(Path(path).read_text())

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def artifact(path):
    path = Path(path)
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': sha(path)}

def verify():
    seal = read(OLD / 'seal.json')['files']
    original = subprocess.check_output(['git', 'show', '2741474cd5:docs/performance/results/change-0811/seal.json'], cwd=ROOT)
    assert original == (OLD / 'seal.json').read_bytes()
    for name, digest in seal.items():
        if name.startswith('docs/performance/results/change-0811/'):
            assert sha(ROOT / name) == digest, name
    source = read(OLD / 'build/source.json')['files']
    names = subprocess.check_output(['git', 'ls-files', '-z', '--', 'crates', 'Cargo.toml', 'clippy.toml', '.cargo/config.toml', 'rust-toolchain.toml'], cwd=ROOT).decode().split('\0')
    assert set(filter(None, names)) == set(source) and len(source) == 9196
    for name, digest in source.items():
        assert sha(ROOT / name) == digest, name
    frozen = read(OLD / 'build/frozen-inputs.json')
    for key in ('architecture', 'unrelated', 'root_inputs'):
        for name, digest in frozen[key].items():
            assert sha(ROOT / name) == digest, name
    assert read(P / 'architecture-inputs.json') == frozen['architecture']
    return {'sealed_packet': artifact(OLD / 'seal.json'), 'source': artifact(OLD / 'build/source.json'),
            'production_files': len(source), 'architecture_files': len(frozen['architecture'])}

if __name__ == '__main__':
    print(json.dumps(verify(), sort_keys=True))
