"""Remove only the completed 0907 source/build directories, after evidence checks."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
S = Path('/tmp/litchi-0907-src')
T = ROOT / 'target/litchi-0907-build'
receipt = P / 'cleanup.json'
assert not receipt.exists()
manifest = json.loads((P / 'source.json').read_text())
for name in ['analysis.json', 'replay-before-cleanup.json']:
    assert json.loads((P / name).read_text())['status'] == 'pass', name
for directory in [S, T]:
    assert directory.is_dir() and not directory.is_symlink()
    assert (directory / '.owned-0907').read_text() == manifest['commit'] + '\n'
for row in manifest['files']:
    path = S / row['path']
    assert path.stat().st_size == row['bytes'] and hashlib.sha256(path.read_bytes()).hexdigest() == row['sha256'], row['path']
binary = json.loads((P / 'binary.json').read_text())
assert hashlib.sha256(Path(binary['path']).read_bytes()).hexdigest() == binary['sha256']
active = []
for proc in Path('/proc').iterdir():
    if not proc.name.isdigit() or int(proc.name) == os.getpid():
        continue
    try:
        args = (proc / 'cmdline').read_bytes().split(b'\0')
        cwd = (proc / 'cwd').resolve()
        if any(cwd == directory or directory in cwd.parents or any(arg.decode(errors='replace').startswith(str(directory) + '/') for arg in args) for directory in [S, T]):
            active.append(proc.name)
    except (OSError, RuntimeError):
        continue
assert not active, active
retained_dir = P / 'binaries'
retained_dir.mkdir(exist_ok=True)
retained = retained_dir / 'native'
assert not retained.exists()
shutil.copy2(binary['path'], retained)
assert hashlib.sha256(retained.read_bytes()).hexdigest() == binary['sha256']
removed = []
for directory in [S, T]:
    files = [f for f in directory.rglob('*') if f.is_file() and not f.is_symlink()]
    removed.append({'path': str(directory), 'regular_files': len(files), 'logical_bytes': sum(f.stat().st_size for f in files)})
    shutil.rmtree(directory)
    assert not directory.exists()
with receipt.open('x') as out:
    json.dump({'status': 'pass', 'completed_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'removed': removed, 'worktrees_created': 0, 'binary_retained': {'path': str(retained), 'bytes': retained.stat().st_size, 'sha256': binary['sha256'], 'git_tracked': False}, 'scope': 'Owned marker-checked source and build roots only; unrelated worktrees untouched.'}, out, indent=2)
    out.write('\n')
print('cleaned', sum(row['logical_bytes'] for row in removed), 'logical bytes')
