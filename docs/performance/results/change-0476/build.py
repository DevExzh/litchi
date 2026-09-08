#!/usr/bin/env python3
"""Build both candidate instrumentation modes from one bound clean checkout."""
from pathlib import Path
import shutil
import subprocess
from common import ROOT, REPO, TEMP, ENV, ENV_KEYS, sha, meta, read, write, now, check_source


def main():
    source = read(ROOT / 'candidate-source.json')
    protocol = read(ROOT / 'protocol.json')
    assert protocol['drivers']['build.py'] == sha(Path(__file__))
    assert protocol['drivers']['common.py'] == sha(ROOT / 'common.py')
    check_source(source)
    argv = ['cargo', 'build', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
        '--target-dir', str(REPO / 'tools/perf-baseline/target'), '--features', 'allocator-metrics',
        '--bin', 'litchi-perf-baseline', '--bin', 'litchi-perf-baseline-alloc']
    record = dict(argv=argv, cwd=source['build_path'], environment={k: ENV[k] for k in ENV_KEYS},
        driver_sha256=sha(Path(__file__)), common_sha256=sha(ROOT / 'common.py'),
        protocol_sha256=sha(ROOT / 'protocol.json'), source_binding_sha256=sha(ROOT / 'candidate-source.json'), started_utc=now())
    write(ROOT / 'build-command.started.json', record)
    with (ROOT / 'build-command.stdout').open('xb') as out, (ROOT / 'build-command.stderr').open('xb') as err:
        result = subprocess.run(argv, cwd=source['build_path'], env=ENV, stdout=out, stderr=err)
    record.update(exit_code=result.returncode, finished_utc=now(),
        artifacts={name: meta(ROOT / name) for name in ('build-command.stdout', 'build-command.stderr')})
    write(ROOT / 'build-command.json', record)
    assert result.returncode == 0
    check_source(source)
    binaries = {}
    for mode, name in [('normal', 'litchi-perf-baseline'), ('allocator', 'litchi-perf-baseline-alloc')]:
        target = TEMP / f'candidate-{mode}'; assert not target.exists()
        shutil.copy2(REPO / 'tools/perf-baseline/target/release' / name, target)
        binaries[mode] = dict(path=str(target), **meta(target))
    write(ROOT / 'candidate-build.json', dict(**source, fresh_build=True, binaries=binaries,
        receipt=dict(path='build-command.json', sha256=sha(ROOT / 'build-command.json')),
        protocol_sha256=sha(ROOT / 'protocol.json')))
    print('candidate build passed', flush=True)


if __name__ == '__main__':
    main()
