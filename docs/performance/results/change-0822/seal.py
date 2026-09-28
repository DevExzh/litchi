"""Seal the 0822 record and check the exact staged or committed batch change set since the measured base."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
DOCS = ['0822-real-file-pptx-edit-profile.md', 'BASELINE.md', 'CRUD_COVERAGE.md', 'GOAL_AUDIT.md', 'HOTSPOTS.md', 'REPORT.md']

def digest(data):
    return hashlib.sha256(data).hexdigest()

def sha(path):
    return digest(path.read_bytes())

def main():
    assert len(sys.argv) == 2 and sys.argv[1] in ('--write', '--check-index', '--check-head')
    subprocess.run([sys.executable, '-B', str(P / 'validate.py'), '--final'], cwd=ROOT, check=True)
    assert not list(P.rglob('__pycache__'))
    paths = [path for path in P.rglob('*') if path.is_file() and path.name != 'seal.json']
    assert not any(path.is_symlink() for path in paths)
    paths += [ROOT / 'docs/performance' / name for name in DOCS]
    paths += [ROOT / name for name in json.loads((P / 'origin.json').read_text())['tool_allowlist']]
    files = {str(path.relative_to(ROOT)): sha(path) for path in sorted(paths)}
    result = {'schema': 'litchi.performance.0822.seal.v1', 'files': files}
    encoded = json.dumps(result, indent=2, sort_keys=True) + '\n'
    if sys.argv[1] == '--write':
        (P / 'seal.json').write_text(encoded)
    else:
        assert (P / 'seal.json').read_text() == encoded
        expected = files | {str((P / 'seal.json').relative_to(ROOT)): digest(encoded.encode())}
        index = sys.argv[1] == '--check-index'
        base = json.loads((P / 'origin.json').read_text())['base']
        command = ['git', 'diff', '--cached', '--name-only', '-z', base] if index else ['git', 'diff', '--name-only', '-z', base, 'HEAD']
        actual = {name for name in subprocess.check_output(command, cwd=ROOT).decode().split('\0') if name}
        assert actual == set(expected), (actual - set(expected), set(expected) - actual)
        revision = ':' if index else 'HEAD:'
        for name, wanted in expected.items():
            assert digest(subprocess.check_output(['git', 'show', revision + name], cwd=ROOT)) == wanted, name
    print('0822 seal PASS:', len(files) + 1, 'owned paths including seal')

if __name__ == '__main__':
    main()
