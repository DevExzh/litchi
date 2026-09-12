"""Build and serially capture the separately instrumented publication method."""
import argparse
import csv
import datetime
import json
import os
import shutil
import subprocess
import time

from run import HERE, REPO, SCRATCH, TARGET, sha, sources, write
from capture import validate, prior

PROBE = HERE / 'allocator-probe'


def probe_sources():
    return {str(p.relative_to(REPO)): sha(p) for p in sorted(PROBE.rglob('*'))
            if p.is_file()}


def build(variant):
    directory = HERE / ('allocator-' + variant)
    directory.mkdir()
    manifest = sources()
    probe = probe_sources()
    write(directory / 'source-manifest.json', manifest)
    write(directory / 'probe-manifest.json', probe)
    settings = dict(CARGO_TARGET_DIR=str(TARGET / 'allocator'), CARGO_BUILD_JOBS='2',
                    CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_DEBUG='0', TMPDIR=str(SCRATCH))
    command = ['cargo', 'build', '--release', '--locked', '--manifest-path',
               str(PROBE / 'Cargo.toml'), '--features', 'allocator-metrics']
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    tick = time.monotonic()
    with (directory / 'build.log').open('x') as out:
        result = subprocess.run(['/usr/bin/time', '-v', *command], cwd=REPO,
                                env=dict(os.environ, **settings), stdout=out, stderr=subprocess.STDOUT)
    receipt = dict(command=command, environment=settings, started_utc=started,
                   elapsed_seconds=time.monotonic() - tick, exit_code=result.returncode,
                   source_unchanged=sources() == manifest, probe_unchanged=probe_sources() == probe,
                   source_manifest_sha256=sha(directory / 'source-manifest.json'),
                   probe_manifest_sha256=sha(directory / 'probe-manifest.json'),
                   log_sha256=sha(directory / 'build.log'))
    if result.returncode == 0:
        binary = SCRATCH / ('allocation-' + variant)
        assert not binary.exists()
        shutil.copy2(TARGET / 'allocator/release/litchi-docx-publication-allocation-probe', binary)
        receipt.update(binary=str(binary), binary_sha256=sha(binary))
    write(directory / 'build-receipt.json', receipt)
    assert result.returncode == 0 and receipt['source_unchanged'] and receipt['probe_unchanged']
    print('allocator build', variant, 'passed', flush=True)


def capture(variant, repeat):
    build_dir = HERE / ('allocator-' + variant)
    receipt = json.loads((build_dir / 'build-receipt.json').read_text())
    binary = SCRATCH / ('allocation-' + variant)
    assert sha(binary) == receipt['binary_sha256']
    manifest = json.loads((build_dir / 'source-manifest.json').read_text())
    assert sources() == manifest
    assert probe_sources() == json.loads((build_dir / 'probe-manifest.json').read_text())
    directory = HERE / f'alloc-{variant}-r{repeat}'
    directory.mkdir()
    names = sorted(f'p{p}-k{k}-{provider}-{mode}' for p in [128, 512]
                   for k in [1, 8, 32] for provider in ['owned', 'file'] for mode in ['repeated', 'batch'])
    if repeat == 2:
        names.reverse()
    for name in names:
        p, k, provider, mode = prior.parse_name(name)
        scratch = SCRATCH / 'corpora' / (directory.name + '-' + name)
        assert not scratch.exists()
        output, stdout, stderr = [directory / (name + suffix) for suffix in ['.csv', '.stdout', '.stderr']]
        command = ['/usr/bin/time', '-v', 'taskset', '-c', '2', str(binary),
                   '--paragraphs', str(p), '--replacements', str(k), '--source', provider,
                   '--mode', mode, '--samples', '1', '--warmups', '0', '--repeats', '1',
                   '--artifact-dir', str(scratch), '--output', str(output)]
        started = datetime.datetime.now(datetime.timezone.utc).isoformat()
        tick = time.monotonic()
        with stdout.open('x') as out, stderr.open('x') as err:
            result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err)
        record = dict(command=command, started_utc=started, elapsed_seconds=time.monotonic() - tick,
                      exit_code=result.returncode, source_unchanged=sources() == manifest,
                      binary_sha256=sha(binary), source_manifest_sha256=sha(build_dir / 'source-manifest.json'),
                      probe_manifest_sha256=sha(build_dir / 'probe-manifest.json'),
                      candidate_plan_sha256=sha(HERE / 'candidate-plan.json'),
                      cleanup_verified=not scratch.exists(),
                      artifacts={p.name: sha(p) for p in [output, stdout, stderr] if p.exists()})
        write(directory / (name + '.json'), record)
        assert result.returncode == 0 and record['source_unchanged'] and record['cleanup_verified']
        with output.open(newline='') as stream:
            validate(name, list(csv.DictReader(stream)), 1, 0, 1)
        samples = [json.loads(line) for line in stdout.read_text().splitlines() if line.startswith('{')]
        assert len(samples) == 1
        sample = samples[0]
        assert sample['tag'] == 'allocationSample' and sample['case'] == name
        assert (sample['repeat'], sample['ordinal'], sample['warmup']) == (0, 0, False)
        counts = sample['allocationSample']
        assert counts['status'] == 'measured' and counts['failed_allocation_calls'] == 0
        assert counts['region_peak_live_bytes'] >= max(counts['live_bytes_before'], counts['live_bytes_after'])
        print(directory.name, name, 'passed', flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['build', 'capture'])
    parser.add_argument('variant', choices=['baseline', 'candidate'])
    parser.add_argument('--repeat', type=int, choices=[1, 2], default=1)
    args = parser.parse_args()
    if args.action == 'build':
        build(args.variant)
    else:
        capture(args.variant, args.repeat)
