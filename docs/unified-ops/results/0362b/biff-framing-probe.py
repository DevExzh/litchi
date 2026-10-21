#!/usr/bin/env python3
"""Build/run the isolated BIFF probe; output must be outside the checkout.

Build: python3 docs/unified-ops/results/0362b/biff-framing-probe.py build /external/biff
Run only after other workers acknowledge a quiet interval:
  python3 docs/unified-ops/results/0362b/biff-framing-probe.py run /external/biff
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import statistics
import subprocess

BASE = '90b586ca78'
ROOT = Path(__file__).resolve().parents[4]
SOURCE = Path(__file__).with_suffix('.rs')


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build(output):
    for label in ('before', 'after'):
        project = output / label
        (project / 'src').mkdir(parents=True, exist_ok=True)
        (project / 'biff/src').mkdir(parents=True, exist_ok=True)
        for source in (ROOT / 'crates/litchi-biff/src').glob('*.rs'):
            content = subprocess.check_output(
                ['git', 'show', f'{BASE}:crates/litchi-biff/src/{source.name}'], cwd=ROOT,
            ) if label == 'before' else source.read_bytes()
            (project / 'biff/src' / source.name).write_bytes(content)
        (project / 'biff/Cargo.toml').write_text(
            '[package]\nname="litchi-biff"\nversion="0.0.1"\nedition="2024"\n'
            '[dependencies]\nthiserror="2"\n')
        (project / 'Cargo.toml').write_text(
            '[package]\nname="biff-framing-probe"\nversion="0.0.1"\nedition="2024"\n'
            '[dependencies]\nlitchi-biff={path="biff"}\n'
            '[profile.release]\nlto="fat"\ncodegen-units=1\npanic="abort"\n')
        source = SOURCE.read_text()
        if label == 'before':
            for old, new in (
                ('Records::new(black_box(bytes), Limits::default()).expect("valid limits")',
                 'Records::new(black_box(bytes))'),
                ('Records::new(black_box(bytes), black_box(Limits::default()))',
                 'Records::with_limits(black_box(bytes), black_box(Limits::default()))'),
                ('Encoder::new(Limits::default()).expect("valid limits")', 'Encoder::new()'),
                ('Encoder::new(black_box(Limits::default()))',
                 'Encoder::with_limits(black_box(Limits::default()))'),
            ):
                assert source.count(old) == 1, old
                source = source.replace(old, new)
        (project / 'src/main.rs').write_text(source)
        env = dict(os.environ, CARGO_TARGET_DIR=str(project / 'target'), CARGO_BUILD_JOBS='1')
        with (project / 'build.log').open('w') as log:
            subprocess.run(['cargo', 'build', '--release', '--offline', '--manifest-path',
                            str(project / 'Cargo.toml')], env=env, stdout=log,
                           stderr=subprocess.STDOUT, check=True)


def run(output):
    cpu = min(os.sched_getaffinity(0))
    binaries = {label: output / label / 'target/release/biff-framing-probe'
                for label in ('before', 'after')}
    rows = []
    for route in ('scan-default', 'scan-configured', 'encode-default', 'encode-configured'):
        iterations = 15000 if route.startswith('scan-') else 1000000
        samples = []
        ratios = []
        for block in range(3):
            times = {'before': [], 'after': []}
            for label in ('before', 'after', 'after', 'before'):
                result = json.loads(subprocess.check_output(
                    ['taskset', '-c', str(cpu), str(binaries[label]), route, str(iterations)]))
                result.update(build=label, block=block)
                samples.append(result)
                times[label].append(result['ns'] / iterations)
            ratios.append(statistics.mean(times['after']) / statistics.mean(times['before']))
        assert len({s['checksum'] for s in samples}) == 1
        rows.append(dict(route=route, after_before=ratios, samples=samples))
    result = dict(baseline_commit=BASE, cpu=cpu, source_sha256=digest(SOURCE),
                  rustc=subprocess.check_output(['rustc', '-Vv'], text=True),
                  binaries={label: dict(bytes=p.stat().st_size, sha256=digest(p),
                                         lock_sha256=digest(output / label / 'Cargo.lock'))
                            for label, p in binaries.items()}, rows=rows)
    (output / 'measurements.json').write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps({row['route']: row['after_before'] for row in rows}))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('phase', choices=('build', 'run'))
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    output = args.output.resolve()
    if output.is_relative_to(ROOT):
        parser.error('keep generated crates, targets and evidence outside the checkout')
    (build if args.phase == 'build' else run)(output)
