#!/usr/bin/env python3
"""Check the final checkout, embedded templates and all retained executables."""
import subprocess

from common import ROOT, REPO, TEMP, read, write, meta, sha, now
import gate
import verify


def main():
    protocol, protocol_hash = verify.check_protocol(ROOT)
    binaries = verify.check_binaries(ROOT, protocol)
    source = gate.snapshot()
    assert all(source['sha256'] == spec['source_manifest_sha256']
               for spec in binaries.values())
    checked_binaries = {}
    preliminary_cleanup = read(ROOT / 'preliminary-runtime-cleanup.json')
    assert preliminary_cleanup['status'] == 'pass'
    assert preliminary_cleanup['originals_absent']
    earlier = ROOT / 'preliminary/formal-before-opc-guard'
    first = earlier / 'preliminary/formal-before-cleanup-fix'
    for role, bundle, directory in (
        ('final', ROOT, TEMP),
        ('before-opc-guard', earlier, TEMP / 'preliminary-opc-guard'),
        ('before-cleanup', first, TEMP / 'preliminary'),
    ):
        for instrumentation, spec in read(bundle / 'binaries.json').items():
            path = directory / instrumentation / 'pptx_metadata_spool'
            identity = {key: spec[key] for key in ('bytes', 'sha256')}
            key = f'{role}-{instrumentation}'
            if role == 'final':
                assert meta(path) == identity
            else:
                assert not path.exists()
                assert preliminary_cleanup['binaries'][key] == dict(
                    path=str(path), **identity)
            checked_binaries[key] = dict(path=str(path), **identity,
                                        live=role == 'final')
    inputs = read(ROOT / 'embedded-inputs.json')['files']
    for name, expected in inputs.items():
        identity = {key: expected[key] for key in ('bytes', 'sha256')}
        assert meta(REPO / name) == meta(ROOT / expected['path']) == identity
    reconstructed = {}
    for bundle, source_path, retained in (
        (first, 'tools/perf-baseline/src/pptx_metadata_spool.rs',
         'cleanup-before-pptx-source.txt'),
        (earlier, 'crates/litchi-opc/src/phys_pkg.rs',
         'opc-before-percent-guard-source.txt'),
    ):
        build = read(bundle / 'normal-build.json')
        manifest = read(bundle / build['source_after']['path'])
        assert sha(ROOT / retained) == manifest[source_path]
        reconstructed[retained] = manifest[source_path]
    for patch in ('harness-cleanup.patch', 'opc-percent-guard.patch'):
        subprocess.run(['git', 'apply', '--reverse', '--check', str(ROOT / patch)],
                       cwd=REPO, check=True)
    reports = verify.check_capture_receipts(ROOT, protocol, protocol_hash, binaries)
    verify.check_summary(ROOT, reports)
    validation = verify.check_rust_validation(ROOT, binaries, protocol)
    preliminary = verify.check_preliminary(ROOT, protocol)
    result = dict(
        schema='pptx-metadata-spool-live-state-v1', checked_utc=now(),
        revision=subprocess.check_output(['git', 'rev-parse', 'HEAD'],
                                         cwd=REPO, text=True).strip(),
        source=source, embedded_input_count=len(inputs), binaries=checked_binaries,
        reconstructed_sources=reconstructed,
        validation=validation, preliminary=preliminary, status='pass',
        protected_files={name: sha(REPO / name) for name in
                         ('docs/GOAL.md', 'docs/report/spec-gap-audit.md')})
    write(ROOT / 'live-state.json', result)
    print('Live source, 19 templates, two final binaries and four prior cleanup identities verified.')


if __name__ == '__main__':
    main()
