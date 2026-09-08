#!/usr/bin/env python3
"""Build and copy separately instrumented diagnostic executables under cpu.lock."""
import shutil
import re
from pathlib import Path
import subprocess
import sys
from common import ROOT, REPO, TEMP, ENV, read, write, meta, sha, now


def embedded_inputs():
    owner = REPO / 'crates/litchi-docx/src'
    names = re.findall(r'include_str!\("([^"\n]+)"\)', (owner / 'template.rs').read_text())
    assert len(names) == len(set(names)) == 8
    inputs = [(owner / name, Path(name)) for name in sorted(names)]
    files = {}
    for source, relative in inputs:
        assert not relative.is_absolute() and '..' not in relative.parts
        destination = ROOT / 'embedded-resources' / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        before = meta(source)
        if not destination.exists():
            shutil.copy2(source, destination)
        assert meta(destination) == before == meta(source)
        files[source.relative_to(REPO).as_posix()] = dict(
            path=destination.relative_to(ROOT).as_posix(), **before)
    value = dict(schema='docx-embedded-inputs-v1', files=files)
    manifest = ROOT / 'embedded-inputs.json'
    if manifest.exists():
        assert read(manifest) == value
    else:
        write(manifest, value)
    return dict(path=manifest.name, **meta(manifest))


def main():
    arm = sys.argv[1]
    assert arm in ('control', 'candidate')
    retained_inputs = embedded_inputs()
    binaries = {}
    for instrumentation in ('normal', 'allocator'):
        label = f'build-{arm}-{instrumentation}'
        command = ['cargo', 'build', '--release', '--locked', '--manifest-path',
                   'tools/perf-baseline/Cargo.toml', '--bin', 'docx_plain_paragraph_tail_append']
        if instrumentation == 'allocator':
            command += ['--features', 'allocator-metrics']
        assert embedded_inputs() == retained_inputs
        subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'), label, *command],
                       cwd=REPO, env=ENV, check=True)
        assert embedded_inputs() == retained_inputs
        gate_path = ROOT / 'validation' / f'{label}.json'
        gate = read(gate_path)
        assert gate['exit_code'] == 0 and gate['source_unchanged']
        origin = REPO / 'tools/perf-baseline/target/release/docx_plain_paragraph_tail_append'
        destination = TEMP / f'{arm}-{instrumentation}' / 'docx_plain_paragraph_tail_append'
        destination.parent.mkdir(exist_ok=True)
        assert not destination.exists()
        origin_meta = meta(origin)
        shutil.copy2(origin, destination)
        assert meta(destination) == origin_meta == meta(origin)
        build = dict(gate, embedded_inputs=retained_inputs,
                     embedded_inputs_unchanged=True, binary=dict(path=str(destination), **origin_meta),
                     original_binary=dict(path=str(origin), **origin_meta),
                     copied_utc=now(),
                     validation_gate=dict(path=gate_path.relative_to(ROOT).as_posix(),
                                          sha256=sha(gate_path)))
        build_path = ROOT / f'{arm}-{instrumentation}-build.json'
        write(build_path, build)
        binaries[instrumentation] = dict(
            path=str(destination), **origin_meta, embedded_inputs=retained_inputs,
            embedded_inputs_unchanged=True,
            build_gate=build_path.name, build_gate_sha256=sha(build_path),
            source_manifest_sha256=gate['source_after']['sha256'])
    assert binaries['normal']['source_manifest_sha256'] == binaries['allocator']['source_manifest_sha256']
    write(ROOT / f'{arm}-binaries.json', binaries)


if __name__ == '__main__':
    main()
