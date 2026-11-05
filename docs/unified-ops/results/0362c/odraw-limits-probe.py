#!/usr/bin/env python3
"""Build outside checkout; run only during a coordinator-approved quiet window."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import statistics
import subprocess

BASE = '45a5f9e9d1'
ROOT = Path(__file__).resolve().parents[4]
SOURCE = Path(__file__).with_suffix('.rs')
ROUTES = ('blip-small', 'blip-large', 'delayed', 'shape-parse', 'shape-record', 'recursive')

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def build(output):
    for label in ('before', 'after'):
        project = output / label
        (project / 'src').mkdir(parents=True, exist_ok=True)
        sources = {}
        for source in sorted((ROOT / 'crates/litchi-odraw/src').rglob('*.rs')):
            rel = source.relative_to(ROOT / 'crates/litchi-odraw')
            content = subprocess.check_output(['git', 'show', f'{BASE}:{source.relative_to(ROOT)}'], cwd=ROOT) if label == 'before' else source.read_bytes()
            target = project / 'odraw' / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(content)
            sources[str(source.relative_to(ROOT))] = digest(target)
        (project / 'source-digests.json').write_text(json.dumps(sources, indent=2)+'\n')
        (project / 'odraw/Cargo.toml').write_text('[package]\nname="litchi-odraw"\nversion="0.0.1"\nedition="2024"\n[dependencies]\nbitflags={version="=2.13.2",features=["std","serde"]}\nzerocopy={version="=0.8.62",features=["std"]}\nzerocopy-derive="=0.8.62"\n')
        (project / 'Cargo.toml').write_text('[package]\nname="odraw-limits-probe"\nversion="0.0.1"\nedition="2024"\n[dependencies]\nlitchi-odraw={path="odraw"}\n[profile.release]\nlto="fat"\ncodegen-units=1\npanic="abort"\n')
        code = SOURCE.read_text()
        if label == 'before':
            for old, new in [('image::Blip::parse(', 'image::Blip::parse_with('), ('image::Store::parse(', 'image::Store::parse_with('), ('image::Delay::new(', 'image::Delay::with_limits('), ('shape::parse(', 'shape::parse_with('), ('.find_recursive(', '.find_recursive_with('), ('shape::Shape::from_record(black_box(shape_record.clone()), Limits::default())', 'shape::Shape::try_from(black_box(shape_record.clone()))')]:
                assert old in code, old
                code = code.replace(old, new)
        (project / 'src/main.rs').write_text(code)
        env = dict(os.environ, CARGO_TARGET_DIR=str(project / 'target'), CARGO_BUILD_JOBS='1')
        with (project / 'build.log').open('w') as log:
            subprocess.run(['cargo','build','--release','--offline','--manifest-path',str(project/'Cargo.toml')],env=env,stdout=log,stderr=subprocess.STDOUT,check=True)

def run(output, target_ms=250, blocks=3, routes=ROUTES):
    heavy = {'cargo', 'rustc', 'rustdoc', 'rustfmt', 'cargo-fmt', 'clippy-driver', 'heaptrack', 'valgrind'}
    processes = subprocess.check_output(['ps', '-eo', 'comm'], text=True).splitlines()
    competing = [p.strip() for p in processes if p.strip() in heavy or p.strip().startswith(('litchi_', 'xlsx_', 'ppt-digest', 'scoped_patch_', 'odraw-limits', 'candidate-', 'baseline-'))]
    if competing:
        raise RuntimeError(f'quiet window required; competing processes: {competing}')
    cpu = min(os.sched_getaffinity(0))
    binaries = {label:output/label/'target/release/odraw-limits-probe' for label in ('before','after')}
    def sample(label, route, iterations):
        return json.loads(subprocess.check_output(['taskset','-c',str(cpu),str(binaries[label]),route,str(iterations)]))
    rows=[]
    for route in routes:
        calibration=sample('before',route,10000)
        iterations=max(10000, int(target_ms * 1_000_000 / max(1, calibration['ns']) * 10000))
        samples=[]
        ratios=[]
        for block in range(blocks):
            times={'before':[], 'after':[]}
            for label in ('before','after','after','before'):
                result=sample(label,route,iterations)
                result.update(build=label,block=block)
                samples.append(result)
                times[label].append(result['ns']/iterations)
            ratios.append(statistics.mean(times['after'])/statistics.mean(times['before']))
        assert len({s['checksum'] for s in samples}) == 1
        row=dict(route=route,calibration=calibration,after_before=ratios,samples=samples)
        rows.append(row)
        (output/f'route-{route}.json').write_text(json.dumps(row,indent=2)+'\n')
        print(json.dumps({route:ratios}),flush=True)
    result=dict(baseline_commit=BASE,cpu=cpu,target_ms=target_ms,blocks=blocks,source_sha256=digest(SOURCE),runner_sha256=digest(Path(__file__)),rustc=subprocess.check_output(['rustc','-Vv'],text=True),binaries={label:dict(bytes=p.stat().st_size,sha256=digest(p),lock_sha256=digest(output/label/'Cargo.lock'),source_digests=json.loads((output/label/'source-digests.json').read_text())) for label,p in binaries.items()},rows=rows)
    (output/'measurements.json').write_text(json.dumps(result,indent=2)+'\n')
    print(json.dumps({r['route']:r['after_before'] for r in rows}))

if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('phase',choices=('build','run'))
    parser.add_argument('output',type=Path)
    parser.add_argument('--target-ms',type=int,default=250)
    parser.add_argument('--blocks',type=int,default=3)
    parser.add_argument('--route',choices=ROUTES,action='append')
    args=parser.parse_args()
    output=args.output.resolve()
    if output.is_relative_to(ROOT): parser.error('keep build artifacts outside checkout')
    if args.target_ms <= 0 or args.blocks <= 0: parser.error('positive target duration and block count required')
    if args.phase == 'build': build(output)
    else: run(output, args.target_ms, args.blocks, args.route or ROUTES)
