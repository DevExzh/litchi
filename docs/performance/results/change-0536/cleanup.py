"""Remove only the frozen campaign's owned scratch after checking live references."""
import datetime
import json
import os
from pathlib import Path
import shutil
from run import HERE, sha, write

plan = json.loads((HERE / 'plan.json').read_text())
owned = [Path(name) for name in plan['owned_paths']]
assert owned == [Path('/tmp/litchi-goal-0536'), Path('/home/zhuhe/litchi-goal-0536-target')]
assert not owned[1].is_symlink()
if owned[0].is_symlink():
    assert owned[0].resolve() == owned[1] / 'retained-binaries'


def belongs(value):
    return any(value == str(path) or value.startswith(str(path) + '/') for path in owned)


references = []
for proc in Path('/proc').iterdir():
    if not proc.name.isdigit() or int(proc.name) == os.getpid():
        continue
    for name in ('cwd', 'exe'):
        try:
            value = os.readlink(proc / name).removesuffix(' (deleted)')
        except OSError:
            continue
        if belongs(value):
            references.append(dict(pid=int(proc.name), kind=name, path=value))
    try:
        args = (proc / 'cmdline').read_bytes().split(b'\0')
        for raw in args:
            value = raw.decode(errors='replace')
            if belongs(value):
                references.append(dict(pid=int(proc.name), kind='argument', path=value))
        for line in (proc / 'maps').read_text().splitlines():
            fields = line.split(maxsplit=5)
            if len(fields) == 6 and belongs(fields[5].removesuffix(' (deleted)')):
                references.append(dict(pid=int(proc.name), kind='mapping', path=fields[5]))
    except (OSError, UnicodeError):
        pass
    try:
        for fd in (proc / 'fd').iterdir():
            try:
                value = os.readlink(fd).removesuffix(' (deleted)')
            except OSError:
                continue
            if belongs(value):
                references.append(dict(pid=int(proc.name), kind='fd', path=value))
    except OSError:
        pass
assert not references, references
for path in owned:
    if path.is_symlink():
        path.unlink()
    elif path.exists():
        shutil.rmtree(path)
for path in HERE.rglob('__pycache__'):
    assert not path.is_symlink()
    shutil.rmtree(path)
assert all(not path.exists() for path in owned)
assert not list(HERE.rglob('__pycache__'))
write(HERE / 'cleanup.json', dict(
    observed_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
    plan_sha256=sha(HERE / 'plan.json'), removed=[str(path) for path in owned],
    accessible_process_references=references, owned_paths_absent=True,
    python_cache_absent=True,
    scope='Accessible process cwd/exe, arguments, mappings, and file descriptors checked before removal.',
))
print('Removed both owned campaign trees; no accessible live references or Python cache remain.')
