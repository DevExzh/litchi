"""Remove only marked 0910 source/build roots after verified evidence replay."""
import datetime
import hashlib
import json
from pathlib import Path
import shutil

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PARENT = '468fe2931d6532c32f4c6608590816b8eaa88e31'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    assert not (P / 'cleanup.json').exists()
    assert json.loads((P / 'replay-before-cleanup.json').read_text())['status'] == 'pass'
    source = json.loads((P / 'source-v1.json').read_text())
    base = json.loads((P / 'source-base.json').read_text())
    roots = [Path(base[key]) for key in ('source_root', 'build_root')]
    assert roots == [Path('/tmp/litchi-0910-src'), ROOT / 'target/litchi-0910-build']
    for d in roots:
        assert d.is_dir() and not d.is_symlink() and (d / '.owned-0910').read_text() == PARENT + '\n'
    active = []
    for proc in Path('/proc').iterdir():
        if not proc.name.isdigit():
            continue
        try:
            cwd = (proc / 'cwd').resolve(strict=True)
        except (FileNotFoundError, PermissionError, ProcessLookupError):
            continue
        if any(cwd.is_relative_to(d) for d in roots):
            active.append({'pid': proc.name, 'cwd': str(cwd)})
    assert not active, active
    for row in source['files']:
        f = roots[0] / row['path']
        assert f.stat().st_size == row['bytes'] and sha(f) == row['sha256'], f
    relocation = json.loads((P / 'binary.json').read_text())
    retained = ROOT / relocation['retained']['path']
    original = Path(relocation['original']['path'])
    if not original.is_absolute():
        original = ROOT / original
    assert sha(retained) == sha(original) == relocation['retained']['sha256'] == relocation['original']['sha256']
    removed = []
    for d in roots:
        files = [f for f in d.rglob('*') if f.is_file()]
        removed.append({'path': str(d), 'files': len(files), 'logical_bytes': sum(f.stat().st_size for f in files)})
        shutil.rmtree(d)
        assert not d.exists()
    result = {'status': 'pass', 'utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'removed': removed, 'active_cwd_processes': active, 'source_inventory_verified': len(source['files']), 'worktrees_created': 0, 'unrelated_worktrees_untouched': True, 'retained_evidence': 'Native binary and generated archive evidence remain Gitignored under this packet, with tracked descriptors.'}
    with (P / 'cleanup.json').open('x') as stream:
        json.dump(result, stream, indent=2)
        stream.write('\n')
    print('0910 cleanup PASS:', sum(r['logical_bytes'] for r in removed), 'logical bytes removed')


if __name__ == '__main__':
    main()
