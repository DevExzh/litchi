"""Serial source-bound quality checks; each attempt keeps its exact test patch."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
TARGET = Path('/home/zhuhe/litchi-goal-0541-target')

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def write(path, value):
    path.write_text(json.dumps(value, indent=2) + '\n')

def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()

def manifest():
    roots = ['crates', 'tools/perf-baseline', 'Cargo.toml', 'Cargo.lock', '.cargo', 'rust-toolchain.toml']
    tracked = subprocess.check_output(['git','ls-files','-z','--',*roots],cwd=REPO).split(b'\0')
    other = subprocess.check_output(['git','ls-files','--others','--exclude-standard','-z','--',*roots],cwd=REPO).split(b'\0')
    names = {n.decode() for n in tracked if n}
    names.update(n.decode() for n in other if n and n.endswith(b'.rs'))
    return {name:sha(REPO/name) for name in sorted(names) if (REPO/name).is_file()}

def freeze(stage):
    assert not (HERE/'SHA256SUMS').exists()
    folder = HERE/stage
    folder.mkdir(exist_ok=False)
    write(folder/'source-manifest.json',manifest())
    plan=json.loads((HERE/'plan.json').read_text())
    patch=subprocess.check_output(['git','diff',plan['revision'],'--','crates','tools/perf-baseline'],cwd=REPO)
    for name in plan['allowed_changes']:
        p=REPO/name
        if p.exists() and subprocess.run(['git','ls-files','--error-unmatch',name],cwd=REPO,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL).returncode:
            diff=subprocess.run(['git','diff','--no-index','--','/dev/null',name],cwd=REPO,capture_output=True)
            assert diff.returncode==1
            patch+=diff.stdout
    (folder/'source.patch').write_bytes(patch)
    for name in plan['allowed_changes']:
        if (REPO/name).exists():
            path=folder/'sources'/name
            path.parent.mkdir(parents=True,exist_ok=True)
            path.write_bytes((REPO/name).read_bytes())
    print('Frozen',stage,flush=True)

def run(stage,name):
    assert not (HERE/'SHA256SUMS').exists()
    folder=HERE/stage
    assert not list(folder.glob(name+'.*')),'occupied check'
    frozen=json.loads((folder/'source-manifest.json').read_text())
    assert manifest()==frozen
    command=json.loads((HERE/'quality-plan.json').read_text())['commands'][name]
    tmp=TARGET/'test-tmp';tmp.mkdir(parents=True,exist_ok=True)
    environment=dict(os.environ);environment['TMPDIR']=str(tmp)
    start=now();tick=time.monotonic()
    with (folder/(name+'.stdout')).open('w') as out,(folder/(name+'.stderr')).open('w') as err:
        result=subprocess.run(command,cwd=REPO,env=environment,stdout=out,stderr=err)
    end=now();elapsed=time.monotonic()-tick
    after=manifest()
    receipt=dict(command=command,start_utc=start,end_utc=end,seconds=elapsed,exit_code=result.returncode,
        source_manifest_sha256=sha(folder/'source-manifest.json'),source_unchanged=after==frozen,
        script_sha256=sha(HERE/'run.py'),plan_sha256=sha(HERE/'plan.json'),quality_plan_sha256=sha(HERE/'quality-plan.json'),
        environment={key:environment.get(key) for key in ['TMPDIR','RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS','LD_PRELOAD']},
        artifacts={p.name:sha(p) for p in sorted(folder.glob(name+'.*')) if p.is_file()})
    write(folder/(name+'.receipt.json'),receipt)
    assert after==frozen,'source changed during check'
    assert result.returncode==0,(name,result.returncode)
    print(stage,name,'passed',flush=True)

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('stage');parser.add_argument('action');args=parser.parse_args()
    assert args.stage.replace('-','').isalnum() and args.stage not in ['sources']
    if args.action=='freeze':freeze(args.stage)
    elif args.action=='all':
        for name in json.loads((HERE/'quality-plan.json').read_text())['commands']:
            if name!='focused-tests':run(args.stage,name)
    else:run(args.stage,args.action)
