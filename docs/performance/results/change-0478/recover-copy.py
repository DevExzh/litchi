#!/usr/bin/env python3
"""Complete the allocator copy after authenticating the retained quota failure."""
import shutil

from common import ROOT, REPO, TEMP, read, write, meta, sha, now
import build
import gate


def main():
    failure = read(ROOT / 'build-copy-failure.json')
    cleanup = read(ROOT / 'preliminary-runtime-cleanup.json')
    assert failure['errno'] == 122 and cleanup['status'] == 'pass'
    assert cleanup['originals_absent'] and cleanup['partial_copy_removed']
    label = 'build-allocator-opc-guard'
    gate_path = ROOT / 'validation' / f'{label}.json'
    receipt = read(gate_path)
    assert receipt['exit_code'] == 0 and receipt['source_unchanged']
    assert sha(gate_path) == failure['build_gate']['sha256']
    source_before = gate.snapshot()
    assert source_before == receipt['source_after']
    inputs = build.embedded_inputs()
    origin = REPO / 'tools/perf-baseline/target/release/pptx_metadata_spool'
    destination = TEMP / 'allocator/pptx_metadata_spool'
    assert str(origin) == failure['source']
    assert str(destination) == failure['destination'] and not destination.exists()
    identity = meta(origin)
    with origin.open('rb') as source, destination.open('xb') as target:
        shutil.copyfileobj(source, target)
    shutil.copystat(origin, destination)
    assert meta(origin) == meta(destination) == identity
    assert build.embedded_inputs() == inputs
    source_after = gate.snapshot()
    assert source_after == source_before
    recovery_path = ROOT / 'build-copy-recovery.json'
    write(recovery_path, dict(
        schema='pptx-metadata-spool-build-copy-recovery-v1', completed_utc=now(),
        failure_sha256=sha(ROOT / 'build-copy-failure.json'),
        cleanup_sha256=sha(ROOT / 'preliminary-runtime-cleanup.json'),
        driver_sha256=sha(ROOT / 'recover-copy.py'), source_before=source_before,
        source_after=source_after, source_unchanged=True, embedded_inputs=inputs,
        origin=dict(path=str(origin), **identity),
        destination=dict(path=str(destination), **identity), status='pass'))
    wrapper = dict(receipt, embedded_inputs=inputs, embedded_inputs_unchanged=True,
                   binary=dict(path=str(destination), **identity),
                   original_binary=dict(path=str(origin), **identity),
                   copied_utc=now(),
                   validation_gate=dict(path=f'validation/{label}.json',
                                        sha256=sha(gate_path)),
                   copy_recovery=dict(path=recovery_path.name,
                                      sha256=sha(recovery_path)))
    write(ROOT / 'allocator-build.json', wrapper)
    binaries = {}
    for instrumentation in ('normal', 'allocator'):
        path = ROOT / f'{instrumentation}-build.json'
        item = read(path)
        binaries[instrumentation] = dict(
            item['binary'], embedded_inputs=inputs, embedded_inputs_unchanged=True,
            build_gate=path.name, build_gate_sha256=sha(path),
            source_manifest_sha256=item['source_after']['sha256'])
    assert all(item['source_manifest_sha256'] == source_after['sha256']
               for item in binaries.values())
    write(ROOT / 'binaries.json', binaries)
    print('Recovered allocator copy from the successful unchanged build; both final binaries bound.')


if __name__ == '__main__':
    main()
