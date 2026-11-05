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
import time

ROOT = Path(os.environ.get("LITCHI_PROBE_ROOT", str(Path.cwd()))).resolve()
SOURCE = Path(__file__).with_suffix('.rs')
BASE = 'e63a1fe3534bcaaae2ce367339f7f8275877baef'
ROUTES = ('compress-small','compress-large','compress-noise','decompress-small','decompress-large','decompress-noise','bound-large','framing','copy-small','copy-large','write-small','write-large')

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def build(output):
    if output.exists() and any(output.iterdir()):
        raise RuntimeError('use a fresh build directory; pinned source snapshots must be retained')
    for label in ('before','after'):
        project = output / label
        (project / 'src').mkdir(parents=True,exist_ok=True)
        manifest = (ROOT/'Cargo.toml').read_text().replace('members = ["crates/*"]', 'members = ["crates/litchi-iwa-core", "crates/litchi-iwa-common", "crates/litchi-iwa-protos", "crates/litchi-iwa-archive", "crates/litchi-iwa-cache", "crates/litchi-iwa-package", "crates/litchi-core", "crates/soapberry-zip"]')
        manifest += '\n[package]\nname="snappy-probe"\nversion="0.0.1"\nedition="2024"\n[dependencies]\nlitchi-iwa-core={workspace=true}\nlitchi-iwa-archive={workspace=true}\nsha2={workspace=true}\n[features]\nbefore=[]\nallocations=[]\n'
        (project/'Cargo.toml').write_text(manifest)
        shutil.copy2(ROOT/'Cargo.lock', project/'Cargo.lock')
        files={}
        for crate in ('litchi-iwa-core','litchi-iwa-common','litchi-iwa-protos','litchi-iwa-archive','litchi-iwa-cache','litchi-iwa-package','litchi-core','soapberry-zip'):
            prefix = f'crates/{crate}'
            tracked = subprocess.check_output(['git','ls-tree','-r','--name-only',BASE,'--',prefix],cwd=ROOT,text=True).splitlines() if label=='before' else [str(p.relative_to(ROOT)) for p in (ROOT/prefix).rglob('*') if p.is_file() and (p.suffix in ('.rs','.proto') or p.name=='Cargo.toml')]
            for name in tracked:
                if not (name.endswith(('.rs','.proto')) or name.endswith('/Cargo.toml')): continue
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
            shutil.copy2(project/'target/release/snappy-probe',project/f'snappy-{mode}')

def quiet():
    heavy={'cargo','rustc','rustdoc','rustfmt','cargo-fmt','clippy-driver','heaptrack','valgrind'}
    processes=subprocess.check_output(['ps','-eo','comm'],text=True).splitlines()
    matches=[p.strip() for p in processes if p.strip() in heavy or p.strip().startswith(('litchi_','litchi-perf-','ole_w5_probe','xlsx_','ppt-digest','scoped_patch_','odraw-limits','candidate-','baseline-','snappy-timed','snappy-alloc'))]
    if matches: raise RuntimeError(f'quiet window required: {matches}')

def run(output, results, blocks, target_ms, routes):
    quiet();cpu=min(os.sched_getaffinity(0));rows=[]
    results.mkdir(parents=True, exist_ok=True)
    if any(results.glob('route-*.json')) or (results/'samples.jsonl').exists():
        raise RuntimeError('use a fresh results directory; prior observations must be retained')
    sequence=0
    binaries={label:{mode:output/label/f'snappy-{mode}' for mode in ('timed','allocations')} for label in ('before','after')}
    def sample(label,route,n,mode='timed'):
        nonlocal sequence
        quiet()
        command=['taskset','-c',str(cpu),str(binaries[label][mode]),route,str(n)]
        started_ns=time.time_ns()
        result=subprocess.run(command,text=True,capture_output=True)
        trace=dict(sequence=sequence,route=route,build=label,mode=mode,command=command,started_ns=started_ns,finished_ns=time.time_ns(),exit_code=result.returncode,stdout=result.stdout,stderr=result.stderr)
        with (results/'samples.jsonl').open('a') as stream: stream.write(json.dumps(trace)+'\n')
        sequence+=1
        result.check_returncode()
        try:
            quiet()
        except RuntimeError as error:
            (results/'overlap.json').write_text(json.dumps(dict(sequence=trace['sequence'],error=str(error)),indent=2)+'\n')
            raise
        return json.loads(result.stdout)
    for route in routes:
        calibration={}; iterations={}
        for label in binaries:
            observations=[sample(label,route,101)]
            n=max(101,int(target_ms*1000000*101/max(1,observations[0]['ns'])))|1
            check=sample(label,route,n);observations.append(check)
            if check['ns'] < target_ms*800000:
                n=int(n*target_ms*1100000/max(1,check['ns']))|1
                observations.append(sample(label,route,n))
            calibration[label]=observations;iterations[label]=n
        samples=[];ratios=[]
        for block in range(blocks):
            times={'before':[],'after':[]}
            for label in ('before','after','after','before'):
                n=iterations[label];value=sample(label,route,n);value.update(build=label,block=block);samples.append(value);times[label].append(value['ns']/n)
            ratios.append(statistics.mean(times['after'])/statistics.mean(times['before']))
        allocations={label:sample(label,route,1,'allocations') for label in binaries}
        fidelity=all(len({s[key] for s in samples})==1 for key in ('checksum','bytes','bound'))
        minimum_ns=min(s['ns'] for s in samples)
        row=dict(route=route,calibration=calibration,iterations=iterations,samples=samples,after_before=ratios,allocations=allocations,fidelity=fidelity,minimum_leg_ns=minimum_ns,eligible=fidelity and minimum_ns>=150000000);rows.append(row)
        (results/f'route-{route}.json').write_text(json.dumps(row,indent=2)+'\n');print(json.dumps({route:ratios}),flush=True)
    result=dict(baseline_commit=BASE,cpu=cpu,blocks=blocks,target_ms=target_ms,source_sha256=digest(SOURCE),runner_sha256=digest(Path(__file__)),rustc=subprocess.check_output(['rustc','-Vv'],text=True),locks={label:digest(output/label/'Cargo.lock') for label in binaries},rows=rows,binaries={label:{mode:dict(bytes=p.stat().st_size,sha256=digest(p)) for mode,p in modes.items()} for label,modes in binaries.items()},sources={label:json.loads((output/label/'source-digests.json').read_text()) for label in binaries})
    (results/'measurements.json').write_text(json.dumps(result,indent=2)+'\n')

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('phase',choices=('build','run'));p.add_argument('output',type=Path);p.add_argument('--results-dir',type=Path);p.add_argument('--blocks',type=int,default=3);p.add_argument('--target-ms',type=int,default=500);p.add_argument('--route',action='append',choices=ROUTES);a=p.parse_args();out=a.output.resolve()
    if out.is_relative_to(ROOT): p.error('artifacts must remain outside checkout')
    if a.blocks<1 or a.target_ms<1:p.error('positive block count and duration required')
    results=(a.results_dir or out).resolve()
    if results.is_relative_to(ROOT):p.error('results must remain outside checkout')
    build(out) if a.phase=='build' else run(out,results,a.blocks,a.target_ms,a.route or ROUTES)
