#!/usr/bin/env python3
"""Build paired native Graph probes; run only during an agreed quiet window.

  python3 docs/unified-ops/results/0359c/graph-budget-probe.py build /external/graph
  python3 docs/unified-ops/results/0359c/graph-budget-probe.py run /external/graph

Sources, targets and binaries remain outside the checkout. The same tracked
XLS chart fixture and unchanged dependencies serve both Graph revisions.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import statistics
import subprocess

BASE = 'b6f5225d15'
ROOT = Path(__file__).resolve().parents[4]
SOURCE = Path(__file__).with_suffix('.rs')
FIXTURE = ROOT / 'test-data/ole/xls/WithThreeCharts.xls'


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build(output):
    files = subprocess.check_output(['git', 'ls-files', 'crates/litchi-ograph/src'], cwd=ROOT, text=True).splitlines()
    for label in ('before', 'after'):
        project = output / label
        (project / 'src').mkdir(parents=True, exist_ok=True)
        hashes = {}
        for name in files:
            source = ROOT / name
            content = subprocess.check_output(['git', 'show', f'{BASE}:{name}'], cwd=ROOT) if label == 'before' else source.read_bytes()
            destination = project / 'graph' / source.relative_to(ROOT / 'crates/litchi-ograph')
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(content)
            hashes[name] = hashlib.sha256(content).hexdigest()
        dependencies = ''.join(f'{name}={{path={json.dumps(str(ROOT / "crates" / name))}}}\n' for name in ('litchi-biff', 'litchi-cfb', 'litchi-core'))
        (project / 'graph/Cargo.toml').write_text('[package]\nname="litchi-ograph"\nversion="0.0.1"\nedition="2024"\n[dependencies]\n'+dependencies+'thiserror="2"\n')
        (project / 'Cargo.toml').write_text('[package]\nname="graph-budget-probe"\nversion="0.0.1"\nedition="2024"\n[workspace]\n[dependencies]\nlitchi-ograph={path="graph"}\nlitchi-cfb={path='+json.dumps(str(ROOT / 'crates/litchi-cfb'))+'}\n[profile.release]\nlto="fat"\ncodegen-units=1\npanic="abort"\n')
        (project / 'src/main.rs').write_bytes(SOURCE.read_bytes())
        (project / 'graph-sources.json').write_text(json.dumps(hashes, indent=2)+'\n')
        env = dict(os.environ, CARGO_TARGET_DIR=str(project / 'target'), CARGO_BUILD_JOBS='4')
        with (project / 'build.log').open('w') as log:
            subprocess.run(['cargo', 'build', '--release', '--offline', '--manifest-path', str(project / 'Cargo.toml')], cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT, check=True)


def run(output):
    cpu = min(os.sched_getaffinity(0))
    binaries = {label: output / label / 'target/release/graph-budget-probe' for label in ('before', 'after')}
    rows = []
    for route in ('validate', 'scan', 'parse_encode', 'parse_edit_commit', 'parse_apply'):
        iterations = 20_000_000 if route == 'validate' else 20_000
        samples, ratios = [], []
        for block in range(3):
            times = {'before': [], 'after': []}
            for label in ('before', 'after', 'after', 'before'):
                result = json.loads(subprocess.check_output(['taskset', '-c', str(cpu), str(binaries[label]), route, str(iterations), str(FIXTURE)]))
                result.update(build=label, block=block)
                samples.append(result)
                times[label].append(result['ns'] / iterations)
            ratios.append(statistics.mean(times['after']) / statistics.mean(times['before']))
        assert len({s['checksum'] for s in samples}) == 1
        rows.append(dict(route=route, after_before=ratios, samples=samples))
    result = dict(baseline_commit=BASE, cpu=cpu, source_sha256=digest(SOURCE), fixture=dict(path=str(FIXTURE.relative_to(ROOT)), bytes=FIXTURE.stat().st_size, sha256=digest(FIXTURE)), rustc=subprocess.check_output(['rustc', '-Vv'], text=True), binaries={label:dict(bytes=p.stat().st_size, sha256=digest(p), lock_sha256=digest(output/label/'Cargo.lock'), sources_sha256=digest(output/label/'graph-sources.json')) for label,p in binaries.items()}, rows=rows)
    (output/'measurements.json').write_text(json.dumps(result,indent=2)+'\n')
    print(json.dumps({r['route']:r['after_before'] for r in rows}))


if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('phase',choices=('build','run'))
    parser.add_argument('output',type=Path)
    args=parser.parse_args()
    output=args.output.resolve()
    if output.is_relative_to(ROOT):
        parser.error('keep generated crates, targets and evidence outside the checkout')
    (build if args.phase=='build' else run)(output)
