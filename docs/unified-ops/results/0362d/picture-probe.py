#!/usr/bin/env python3
"""Build isolated pinned sources; measure only in a coordinator-approved quiet window."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import statistics
import subprocess

ROOT = Path(__file__).resolve().parents[4]
SOURCE = Path(__file__).with_suffix('.rs')
BASE = '610e3c1afe'
ROUTES = tuple(f'{operation}-{size}' for operation in ('open','view','metadata','noop','commit','write') for size in ('small','large')) + tuple(f'props-{operation}-{count}' for operation in ('open','get') for count in (4,16,17,256))

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def build(output):
    for label in ('before','after'):
        project = output / label
        (project / 'src').mkdir(parents=True,exist_ok=True)
        manifest = (ROOT/'Cargo.toml').read_text().replace('members = ["crates/*"]', 'members = ["crates/litchi-odraw", "crates/litchi-core"]')
        manifest += '\n[package]\nname="picture-probe"\nversion="0.0.1"\nedition="2024"\n[dependencies]\nlitchi-odraw={workspace=true}\n[features]\nbefore=[]\nallocations=[]\n'
        (project/'Cargo.toml').write_text(manifest)
        shutil.copy2(ROOT/'Cargo.lock', project/'Cargo.lock')
        files={}
        for crate in ('litchi-odraw','litchi-core'):
            prefix = f'crates/{crate}'
            tracked = subprocess.check_output(['git','ls-tree','-r','--name-only',BASE,'--',prefix],cwd=ROOT,text=True).splitlines() if label=='before' else [str(p.relative_to(ROOT)) for p in (ROOT/prefix).rglob('*') if p.is_file() and (p.suffix=='.rs' or p.name=='Cargo.toml')]
            for name in tracked:
                if not (name.endswith('.rs') or name.endswith('/Cargo.toml')): continue
                target=project/name;target.parent.mkdir(parents=True,exist_ok=True)
                target.write_bytes(subprocess.check_output(['git','show',f'{BASE}:{name}'],cwd=ROOT) if label=='before' else (ROOT/name).read_bytes())
                files[name]=digest(target)
        (project/'source-digests.json').write_text(json.dumps(files,indent=2)+'\n')
        shutil.copy2(SOURCE,project/'src/main.rs')
        env=dict(os.environ,CARGO_TARGET_DIR=str(project/'target'),CARGO_BUILD_JOBS='1')
        for mode in ('timed','allocations'):
            features=','.join(x for x in ('before' if label=='before' else '', 'allocations' if mode=='allocations' else '') if x)
            cmd=['cargo','build','--offline','--release','--manifest-path',str(project/'Cargo.toml')]
            if features: cmd += ['--features',features]
            with (project/f'build-{mode}.log').open('w') as log:
                subprocess.run(cmd,env=env,stdout=log,stderr=subprocess.STDOUT,check=True)
            shutil.copy2(project/'target/release/picture-probe',project/f'picture-{mode}')

def quiet():
    heavy={'cargo','rustc','rustdoc','rustfmt','cargo-fmt','clippy-driver','heaptrack','valgrind'}
    processes=subprocess.check_output(['ps','-eo','comm'],text=True).splitlines()
    matches=[p.strip() for p in processes if p.strip() in heavy or p.strip().startswith(('litchi_','xlsx_','ppt-digest','scoped_patch_','odraw-limits','candidate-','baseline-','picture-timed','picture-alloc'))]
    if matches: raise RuntimeError(f'quiet window required: {matches}')

def run(output, blocks, target_ms, routes):
    quiet();cpu=min(os.sched_getaffinity(0));rows=[]
    binaries={label:{mode:output/label/f'picture-{mode}' for mode in ('timed','allocations')} for label in ('before','after')}
    def sample(label,route,n,mode='timed'):
        return json.loads(subprocess.check_output(['taskset','-c',str(cpu),str(binaries[label][mode]),route,str(n)]))
    for route in routes:
        calibration=sample('before',route,101)
        n=max(101,int(target_ms*1000000*101/max(1,calibration['ns'])))|1
        samples=[];ratios=[]
        for block in range(blocks):
            times={'before':[],'after':[]}
            for label in ('before','after','after','before'):
                value=sample(label,route,n);value.update(build=label,block=block);samples.append(value);times[label].append(value['ns']/n)
            ratios.append(statistics.mean(times['after'])/statistics.mean(times['before']))
        assert len({s['checksum'] for s in samples})==1
        allocations={label:sample(label,route,1,'allocations') for label in binaries}
        row=dict(route=route,calibration=calibration,samples=samples,after_before=ratios,allocations=allocations);rows.append(row)
        (output/f'route-{route}.json').write_text(json.dumps(row,indent=2)+'\n');print(json.dumps({route:ratios}),flush=True)
    histories={size:{label:sample(label,f'history-{size}',128,'allocations') for label in binaries} for size in ('small','large')}
    result=dict(baseline_commit=BASE,cpu=cpu,blocks=blocks,target_ms=target_ms,source_sha256=digest(SOURCE),runner_sha256=digest(Path(__file__)),rustc=subprocess.check_output(['rustc','-Vv'],text=True),locks={label:digest(output/label/'Cargo.lock') for label in binaries},rows=rows,histories=histories,binaries={label:{mode:dict(bytes=p.stat().st_size,sha256=digest(p)) for mode,p in modes.items()} for label,modes in binaries.items()},sources={label:json.loads((output/label/'source-digests.json').read_text()) for label in binaries})
    (output/'measurements.json').write_text(json.dumps(result,indent=2)+'\n')

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('phase',choices=('build','run'));p.add_argument('output',type=Path);p.add_argument('--blocks',type=int,default=3);p.add_argument('--target-ms',type=int,default=500);p.add_argument('--route',action='append',choices=ROUTES);a=p.parse_args();out=a.output.resolve()
    if out.is_relative_to(ROOT): p.error('artifacts must remain outside checkout')
    if a.blocks<1 or a.target_ms<1:p.error('positive block count and duration required')
    build(out) if a.phase=='build' else run(out,a.blocks,a.target_ms,a.route or ROUTES)
