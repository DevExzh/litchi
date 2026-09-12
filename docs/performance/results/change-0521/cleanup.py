"""Remove only this campaign's owned build paths after terminal gates."""
import json
import os
from pathlib import Path
import shutil

from checks import COMMANDS
from run import HERE, SCRATCH, TARGET


def main():
    for name, _ in COMMANDS:
        receipt = json.loads((HERE / 'candidate' / ('check-' + name + '.receipt.json')).read_text())
        assert receipt['exit_code'] == 0, name
    owned = [SCRATCH, TARGET]
    assert [str(p) for p in owned] == json.loads((HERE / 'plan.json').read_text())['owned_paths']
    excluded = set()
    pid = os.getpid()
    while pid:
        excluded.add(pid)
        try:
            status = Path('/proc') / str(pid) / 'status'
            parent = next(line for line in status.read_text().splitlines() if line.startswith('PPid:'))
            pid = int(parent.split()[1])
        except (OSError, StopIteration):
            break
    references = []
    for proc in Path('/proc').iterdir():
        if not proc.name.isdigit() or int(proc.name) in excluded:
            continue
        try:
            values = [proc.joinpath('cmdline').read_bytes().decode(errors='replace')]
            for link in [proc / 'cwd', proc / 'exe', *proc.joinpath('fd').iterdir()]:
                try:
                    values.append(os.readlink(link))
                except OSError:
                    pass
            for path in owned:
                if any(str(path) in value for value in values):
                    references.append(dict(pid=int(proc.name), owned_path=str(path)))
        except OSError:
            continue
    assert not references, references
    for path in owned:
        if path.exists():
            shutil.rmtree(path)
    for path in HERE.rglob('__pycache__'):
        shutil.rmtree(path)
    report = dict(removed=[str(p) for p in owned], accessible_process_references=references,
                  owned_paths_absent=all(not p.exists() for p in owned),
                  python_cache_absent=not list(HERE.rglob('__pycache__')),
                  scope='Accessible /proc cmdline, cwd, executable and descriptor links; own ancestors excluded.')
    (HERE / 'cleanup.json').write_text(json.dumps(report, indent=2) + '\n')
    print(json.dumps(report))


if __name__ == '__main__':
    main()
