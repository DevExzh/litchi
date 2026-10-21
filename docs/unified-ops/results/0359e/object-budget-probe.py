#!/usr/bin/env python3
"""Paired OLE object probe. Keep generated projects and captures outside checkout.
Build before requesting a quiet window; run only after every worker acknowledges.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import statistics
import subprocess
import tomllib

BASE = '7d5de8000a'
ROOT = Path(__file__).resolve().parents[4]
SOURCE = Path(__file__).with_suffix('.rs')

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def build(output):
    baseline = subprocess.check_output(['git', 'ls-tree', '-r', '--name-only', BASE, '--', 'crates/litchi-ole-common/src'], cwd=ROOT, text=True).splitlines()
    current = subprocess.check_output(['rg', '--files', 'crates/litchi-ole-common/src'], cwd=ROOT, text=True).splitlines()
    workspace = tomllib.loads((ROOT/'Cargo.toml').read_text())['workspace']['dependencies']
    for label, names in [('before',baseline),('after',current)]:
        project=output/label
        (project/'src').mkdir(parents=True,exist_ok=True)
        hashes={}
        for name in names:
            content = subprocess.check_output(['git','show',f'{BASE}:{name}'],cwd=ROOT) if label=='before' else (ROOT/name).read_bytes()
            destination=project/'common'/Path(name).relative_to('crates/litchi-ole-common')
            destination.parent.mkdir(parents=True,exist_ok=True)
            destination.write_bytes(content)
            hashes[name]=hashlib.sha256(content).hexdigest()
        dependencies={}
        for name in ('chrono','litchi-cfb','litchi-codepage','litchi-core','quick-xml'):
            config=workspace[name]
            if isinstance(config,str): config={'version':config}
            else: config=dict(config)
            if 'path' in config: config['path']=str(ROOT/config['path'])
            dependencies[name]=config
        dependency_toml=''.join(name+'={'+','.join(key+'='+json.dumps(value) for key,value in config.items())+'}\n' for name,config in dependencies.items())
        (project/'common/Cargo.toml').write_text('[package]\nname="litchi-ole-common"\nversion="0.0.1"\nedition="2024"\n[features]\ndefault=[]\nperformance-diagnostics=[]\n[dependencies]\n'+dependency_toml)
        (project/'Cargo.toml').write_text('[package]\nname="object-budget-probe"\nversion="0.0.1"\nedition="2024"\n[workspace]\n[dependencies]\nlitchi-ole-common={path="common"}\nlitchi-cfb={path='+json.dumps(str(ROOT/'crates/litchi-cfb'))+'}\n[profile.release]\nlto="fat"\ncodegen-units=1\npanic="abort"\n')
        (project/'src/main.rs').write_bytes(SOURCE.read_bytes())
        (project/'common-sources.json').write_text(json.dumps(hashes,sort_keys=True,indent=2)+'\n')
        env=dict(os.environ,CARGO_TARGET_DIR=str(project/'target'),CARGO_BUILD_JOBS='4')
        with (project/'build.log').open('w') as log:
            subprocess.run(['cargo','build','--release','--offline','--manifest-path',str(project/'Cargo.toml')],cwd=ROOT,env=env,stdout=log,stderr=subprocess.STDOUT,check=True)

def run(output):
    cpu=min(os.sched_getaffinity(0))
    binaries={label:output/label/'target/release/object-budget-probe' for label in ('before','after')}
    rows=[]
    for opaque,streams in [(0,0),(2*1024*1024,0),(0,256)]:
        for route in ('open','begin_edit','noop_commit','replace_stream','apply_patch'):
            # Calibrate each route to at least 150 ms per leg. Tiny fixed-loop
            # legs can be shorter than 1 ms and exaggerate scheduler noise.
            pilot=json.loads(subprocess.check_output(['taskset','-c',str(cpu),str(binaries['before']),route,'100',str(opaque),str(streams)]))
            iterations=max(100, int(150_000_000 / max(1, pilot['ns']/100)))
            samples=[];ratios=[]
            for block in range(3):
                times={'before':[],'after':[]}
                for label in ('before','after','after','before'):
                    sample=json.loads(subprocess.check_output(['taskset','-c',str(cpu),str(binaries[label]),route,str(iterations),str(opaque),str(streams)]))
                    sample.update(build=label,block=block);samples.append(sample)
                    times[label].append(sample['ns']/iterations)
                ratios.append(statistics.mean(times['after'])/statistics.mean(times['before']))
            assert len({s['checksum'] for s in samples})==1
            rows.append(dict(route=route,opaque_bytes=opaque,extra_streams=streams,after_before=ratios,samples=samples))
    result=dict(baseline_commit=BASE,cpu=cpu,source_sha256=digest(SOURCE),fixture='Identical synthetic CFB: selected Object/Payload, opaque root stream, optional 256 inert sibling streams.',rustc=subprocess.check_output(['rustc','-Vv'],text=True),binaries={label:dict(bytes=p.stat().st_size,sha256=digest(p),lock_sha256=digest(output/label/'Cargo.lock'),sources_sha256=digest(output/label/'common-sources.json')) for label,p in binaries.items()},rows=rows)
    (output/'measurements.json').write_text(json.dumps(result,indent=2)+'\n')
    print(json.dumps({f"{r['route']}/{r['opaque_bytes']}/{r['extra_streams']}":r['after_before'] for r in rows}))

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('phase',choices=('build','run'));parser.add_argument('output',type=Path)
    args=parser.parse_args();output=args.output.resolve()
    if output.is_relative_to(ROOT):parser.error('output must remain outside checkout')
    (build if args.phase=='build' else run)(output)
