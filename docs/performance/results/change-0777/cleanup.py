"""Remove only the three owned 0777 build targets after final custody checks."""
from pathlib import Path
import shutil
import subprocess

import capture as c
import validate


if __name__ == '__main__':
    assert not (c.PACKET / 'cleanup.json').exists()
    validate.validate()
    refs = {'before': c.BASE_REF, 'after': c.json_read(c.CAPTURE / 'complete.json')['candidate']}
    roots = {'before': c.BEFORE, 'after': c.AFTER}
    for leg in roots:
        assert c.source_manifest(roots[leg], refs[leg]) == c.json_read(c.CAPTURE / f'source-{leg}.json')
    assert c.fixture_manifest() == c.json_read(c.CAPTURE / 'fixtures-before.json')
    assert c.package_inventory() == c.json_read(c.CAPTURE / 'package-inventory-before.json')
    origin = c.json_read(c.PACKET / 'origin.json')
    for name, digest in origin['review_edits'].items():
        assert c.sha256(Path(origin['root']) / name) == digest
    for name, digest in origin['main_unrelated'].items():
        assert c.sha256(c.BEFORE / name) == digest
    for leg in c.json_read(c.CAPTURE / 'binaries.json').values():
        for binary in leg.values():
            assert c.sha256(Path(binary['path'])) == binary['sha256']
    targets = []
    for suffix in ('quality', 'release-before', 'release-after'):
        target = Path('/home/zhuhe/code') / f'litchi-target-0777-{suffix}'
        assert target.is_dir() and not target.is_symlink()
        size = int(subprocess.check_output(['du', '-sb', str(target)], text=True).split()[0])
        shutil.rmtree(target)
        assert not target.exists()
        targets.append({'path': str(target), 'bytes': size, 'removed': True})
    c.json_write(c.PACKET / 'cleanup.json', {
        'source_unchanged': True,
        'fixtures_unchanged': True,
        'package_inventory_unchanged': True,
        'executables_verified_before_removal': True,
        'origin_review_edits_unchanged': True,
        'main_unrelated_files_unchanged': True,
        'targets': targets,
        'scope': 'Only three owned build targets. Worktree, copied root lock and reference symlinks are removed after integration and recorded in the report.',
    })
    print(validate.validate())
