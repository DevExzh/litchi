"""Restore the frozen build inputs to new owned directories for reproduction."""
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tarfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
S = Path('/tmp/litchi-0907-src')
T = ROOT / 'target/litchi-0907-build'
assert not S.exists() and not T.exists(), 'refusing existing source or build directory'
manifest = json.loads((P / 'source.json').read_text())
raw = subprocess.check_output(['git', 'archive', manifest['commit'], '--', *manifest['archive_paths']], cwd=ROOT)
assert hashlib.sha256(raw).hexdigest() == manifest['archive_sha256']
S.mkdir()
T.mkdir(parents=True)
with tarfile.open(fileobj=io.BytesIO(raw)) as archive:
    archive.extractall(S, filter='data')
for name, descriptor in manifest['lock_overlays'].items():
    data = (ROOT / descriptor['artifact']).read_bytes()
    assert hashlib.sha256(data).hexdigest() == descriptor['sha256']
    (S / name).write_bytes(data)
for directory in [S, T]:
    (directory / '.owned-0907').write_text(manifest['commit'] + '\n')
actual = [{'path': str(f.relative_to(S)), 'bytes': f.stat().st_size, 'sha256': hashlib.sha256(f.read_bytes()).hexdigest()} for f in sorted(S.rglob('*')) if f.is_file()]
assert actual == manifest['files']
print('Frozen sources restored; use the retained native-build command/environment in a new evidence packet.')
