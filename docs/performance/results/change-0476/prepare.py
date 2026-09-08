#!/usr/bin/env python3
"""Prepare authenticated control reuse or a clean committed candidate source."""
from pathlib import Path
import shutil
import subprocess
import sys
from common import ROOT, REPO, TEMP, sha, meta, now, read, write, check_source


def main(arm):
    assert arm in ('control', 'candidate')
    prior = read(ROOT / 'reuse/build.json')
    started = now()
    if arm == 'control':
        assert not TEMP.exists() and not Path('/tmp/litchi-goal-0474').exists()
        TEMP.mkdir(); (TEMP / 'cpu.lock').touch()
        Path('/tmp/litchi-goal-0474').mkdir()
        tree = Path('/tmp/litchi-goal-0474/tree')
        revision = prior['revision']
        sources = read(ROOT / 'reuse/sources/source.json')
        binaries = {}
        for mode, name in [('normal', 'litchi-perf-baseline'), ('allocator', 'litchi-perf-baseline-alloc')]:
            cached = REPO / 'tools/perf-baseline/target/release' / name
            assert meta(cached) == {k: prior['binaries'][mode][k] for k in ('bytes', 'sha256')}
            binary = TEMP / f'control-{mode}'
            shutil.copy2(cached, binary)
            binaries[mode] = dict(path=str(binary), **meta(binary))
    else:
        tree = TEMP / 'candidate'
        revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO, text=True).strip()
        paths = subprocess.check_output(['git', 'ls-tree', '-r', '--name-only', revision], cwd=REPO, text=True).splitlines()
        sources = {p: sha(REPO / p) for p in paths if Path(p).suffix in ('.rs', '.toml', '.lock')}
        changed = subprocess.check_output(['git', 'diff', 'HEAD', '--name-only'], cwd=REPO, text=True).splitlines()
        assert not any(p in sources for p in changed), 'commit candidate build inputs first'
    assert not tree.exists()
    tracked = subprocess.check_output(['git', 'ls-tree', '-r', '--name-only', revision], cwd=REPO, text=True).splitlines()
    selected = [p for p in tracked if not p.startswith('docs/') or p in sources]
    subprocess.run(['git', 'worktree', 'add', '--detach', '--no-checkout', str(tree), revision], cwd=REPO, check=True)
    subprocess.run(['git', 'sparse-checkout', 'set', '--no-cone', '--stdin'], cwd=tree,
        input=''.join('/' + p + '\n' for p in selected), text=True, check=True)
    subprocess.run(['git', 'reset', '--hard', revision], cwd=tree, check=True)
    for path, digest in prior['fixtures'].items():
        assert sha(REPO / path) == digest
        target = tree / path; target.parent.mkdir(parents=True, exist_ok=True); shutil.copy2(REPO / path, target)
    manifest = ROOT / 'sources' / f'{arm}.json'; write(manifest, sources)
    source = dict(schema='litchi-0476-source-v1', arm=arm, revision=revision, build_path=str(tree),
        source_manifest=dict(path=manifest.relative_to(ROOT).as_posix(), sha256=sha(manifest), files=len(sources)),
        fixtures=prior['fixtures'])
    check_source(source)
    write(ROOT / f'{arm}-source.json', source)
    write(ROOT / f'{arm}-prepare.json', dict(started_utc=started, finished_utc=now(),
        driver_sha256=sha(Path(__file__)), common_sha256=sha(ROOT / 'common.py'),
        source_binding_sha256=sha(ROOT / f'{arm}-source.json'), selected_files=len(selected), clean_after=True))
    if arm == 'control':
        write(ROOT / 'control-build.json', dict(**source, fresh_build=False, binaries=binaries,
            reused_build_sha256=sha(ROOT / 'reuse/build.json'), prepared_utc=now()))
    print('prepared', arm, revision, len(sources), flush=True)


if __name__ == '__main__':
    main(sys.argv[1])
