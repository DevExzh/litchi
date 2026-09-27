"""Verify retained evidence before removing this batch's three owned targets."""
from pathlib import Path
import json, shutil, time
import measure

P = measure.P

if __name__ == '__main__':
    assert (P / 'allocations-0/complete.json').exists()
    assert not (P / 'cleanup.json').exists()
    binaries = json.loads((P / 'measure-1/binaries.json').read_text())
    for leg, root in measure.ROOTS.items():
        expected = json.loads((P / f'measure-1/source-{leg}.json').read_text())
        assert measure.source(root, measure.REFS[leg]) == expected
        assert measure.sha(Path(binaries[leg]['path'])) == binaries[leg]['sha256']
    fixtures = json.loads((P / 'measure-1/fixtures.json').read_text())
    assert all(measure.sha(measure.AFTER / n) == h for n, h in fixtures.items())
    targets = [Path('/home/zhuhe/code/litchi-target-0776-quality'),
               *measure.TARGETS.values()]
    rows = []
    for target in targets:
        assert target.parent == Path('/home/zhuhe/code')
        assert target.name.startswith('litchi-target-0776-')
        assert target.is_dir() and not target.is_symlink()
        size = sum(p.stat().st_size for p in target.rglob('*') if p.is_file())
        shutil.rmtree(target)
        assert not target.exists()
        rows.append({'path': str(target), 'bytes': size, 'removed': True})
    measure.write(P / 'cleanup.json', {
        'time': time.time(), 'executables_verified_before_removal': True,
        'source_unchanged': True, 'fixtures_unchanged': True, 'targets': rows})
