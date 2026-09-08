#!/usr/bin/env python3
"""Build and copy separately instrumented diagnostic executables under cpu.lock."""
import shutil
import subprocess
import sys
from common import ROOT, REPO, TEMP, ENV, read, write, meta, sha, now


def main():
    binaries = {}
    for instrumentation in ('normal', 'allocator'):
        label = f'build-{instrumentation}'
        command = ['cargo', 'build', '--release', '--locked', '--manifest-path',
                   'tools/perf-baseline/Cargo.toml', '--bin', 'zip_directory_spool']
        if instrumentation == 'allocator':
            command += ['--features', 'allocator-metrics']
        subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'), label, *command],
                       cwd=REPO, env=ENV, check=True)
        gate_path = ROOT / 'validation' / f'{label}.json'
        gate = read(gate_path)
        assert gate['exit_code'] == 0 and gate['source_unchanged']
        origin = REPO / 'tools/perf-baseline/target/release/zip_directory_spool'
        destination = TEMP / instrumentation / 'zip_directory_spool'
        destination.parent.mkdir(exist_ok=True)
        assert not destination.exists()
        origin_meta = meta(origin)
        shutil.copy2(origin, destination)
        assert meta(destination) == origin_meta == meta(origin)
        build = dict(gate, binary=dict(path=str(destination), **origin_meta),
                     original_binary=dict(path=str(origin), **origin_meta),
                     copied_utc=now(),
                     validation_gate=dict(path=gate_path.relative_to(ROOT).as_posix(),
                                          sha256=sha(gate_path)))
        build_path = ROOT / f'{instrumentation}-build.json'
        write(build_path, build)
        binaries[instrumentation] = dict(
            path=str(destination), **origin_meta,
            build_gate=build_path.name, build_gate_sha256=sha(build_path),
            source_manifest_sha256=gate['source_after']['sha256'])
    assert binaries['normal']['source_manifest_sha256'] == binaries['allocator']['source_manifest_sha256']
    write(ROOT / 'binaries.json', binaries)


if __name__ == '__main__':
    main()
