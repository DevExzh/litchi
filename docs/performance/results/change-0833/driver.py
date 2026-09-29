"""Root-owned current-source filesystem qualification, with exclusive receipts."""
from pathlib import Path
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BASE = 'eefaca16e39ace3c1e6219aedbac4fe1c0cd060a'
TARGET = ROOT.parent / 'litchi-target-0833'
SCRATCH = ROOT.parent / 'litchi-fs-0833'
BINARY = TARGET / 'release/litchi-perf-baseline'
CASES = ['opc_file_eager_open', 'opc_file_source_open',
         'opc_file_eager_one_part_atomic_save', 'opc_file_source_one_part_atomic_save',
         'pptx_file_eager_open_selected_slide_lifecycle',
         'pptx_file_source_open_selected_slide_lifecycle']
UNRELATED = ['docs/FORMAT_IMPLEMENTATION_REVIEW.md', 'docs/UNIFIED_OPS_API_DESIGN.md', 'matrix-analysis.json']

def sha(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as f:
        for b in iter(lambda: f.read(1048576), b''): h.update(b)
    return h.hexdigest()

def read(path): return json.loads(Path(path).read_text())

def write(path, value):
    with Path(path).open('x') as f:
        json.dump(value, f, indent=2, sort_keys=True)
        f.write('\n')

def output(argv):
    return subprocess.check_output(argv, cwd=ROOT, text=True).strip()

def desc(path):
    path = Path(path)
    assert path.is_file() and not path.is_symlink(), path
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': sha(path)}

def inventory():
    paths = output(['git','ls-files','crates','tools','Cargo.toml','Cargo.lock',
                    '.cargo','rust-toolchain.toml','rustfmt.toml','clippy.toml',
                    'docs/adr','docs/GOAL.md','docs/CRUD_Scenario_Checklist.md']).splitlines()
    return {s: sha(ROOT/s) for s in paths}

def environment():
    e = os.environ.copy()
    e['CARGO_TARGET_DIR'] = str(TARGET)
    e['CARGO_BUILD_JOBS'] = '2'
    e['CARGO_INCREMENTAL'] = '0'
    for k in ('RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS'):
        assert not e.get(k), f'unexpected {k}'
    return e

def prepare():
    assert output(['git','rev-parse','HEAD']) == BASE
    assert 12 in os.sched_getaffinity(0)
    assert not TARGET.exists() and not SCRATCH.exists()
    assert not (P/'prepare.json').exists()
    norms = read(P.parent/'change-0832/architecture-inputs.json')
    assert all(sha(ROOT/k)==v for k,v in norms.items())
    source = inventory()
    env = environment()
    receipt = dict(base=BASE, source=source, normative=norms,
        unrelated={s:sha(ROOT/s) for s in UNRELATED},
        frozen={s:sha(P/s) for s in ['driver.py','design.md','measurement-plan.json']},
        platform=platform.platform(), machine=platform.machine(),
        affinity=sorted(os.sched_getaffinity(0)),
        cpu=Path('/proc/cpuinfo').read_text(), memory=Path('/proc/meminfo').read_text(),
        rustc=output(['rustc','-Vv']), cargo=output(['cargo','-V']),
        python=sys.version, fincore=desc(Path(shutil.which('fincore')).resolve()),
        fincore_version=output(['fincore','--version']),
        filesystem=output(['findmnt','-T',str(ROOT),'-n','-o','FSTYPE,OPTIONS']),
        environment={k:env.get(k) for k in ['PATH','CARGO_TARGET_DIR','CARGO_BUILD_JOBS',
          'CARGO_INCREMENTAL','RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS','LD_PRELOAD','LD_LIBRARY_PATH']},
        started_unix=time.time())
    for path in (TARGET,SCRATCH):
        path.mkdir()
        write(path/'.owner-0833.json', {'base':BASE,'packet':str(P),'root':str(path)})
    write(P/'prepare.json',receipt)
    print('prepare PASS',len(source),'source inputs',flush=True)

def check():
    p=read(P/'prepare.json')
    assert output(['git','rev-parse','HEAD']) == BASE
    assert inventory()==p['source'], 'source changed'
    assert {s:sha(ROOT/s) for s in UNRELATED}==p['unrelated']
    assert all(sha(P/k)==v for k,v in p['frozen'].items()), 'frozen input changed'
    for path in (TARGET,SCRATCH):
        assert not path.is_symlink()
        assert read(path/'.owner-0833.json')=={'base':BASE,'packet':str(P),'root':str(path)}
    e=environment()
    assert {k:e.get(k) for k in p['environment']}==p['environment']
    return p

def run(label, argv):
    check()
    d=P/'commands'/label
    d.mkdir(parents=True, exist_ok=False)
    start=time.time()
    write(d/'started.json',dict(argv=argv,cwd=str(ROOT),started_unix=start,
          prepare_sha256=sha(P/'prepare.json'),driver_sha256=sha(P/'driver.py')))
    code=None; error=None
    with (d/'output.log').open('xb') as log:
        try:
            result=subprocess.run(argv,cwd=ROOT,env=environment(),stdout=log,stderr=subprocess.STDOUT)
            code=result.returncode
        except Exception as exc: error=repr(exc)
    r=dict(argv=argv,started_unix=start,finished_unix=time.time(),exit_code=code,
           error=error,log=desc(d/'output.log'),prepare_sha256=sha(P/'prepare.json'))
    write(d/'receipt.json',r)
    print(label,code,error,flush=True)
    return r

def build():
    r=run('build',['cargo','build','--offline','--locked','--release','--manifest-path',
            'tools/perf-baseline/Cargo.toml','--bin','litchi-perf-baseline'])
    assert r['exit_code']==0, r
    write(P/'build.json',dict(binary=desc(BINARY),receipt=desc(P/'commands/build/receipt.json')))

def quality():
    r=run('quality',['cargo','test','--offline','--locked','--manifest-path',
            'tools/perf-baseline/Cargo.toml','--lib'])
    assert r['exit_code']==0,r
    write(P/'quality.json',dict(receipt=desc(P/'commands/quality/receipt.json'),status='pass'))

def qualify():
    assert desc(BINARY)==read(P/'build.json')['binary']
    rows=[]
    for i,case in enumerate(CASES):
        path=P/f'qualification-{i:02}.json'
        assert not path.exists()
        argv=['taskset','-c','12',str(BINARY),'--case',case,'--warmup','0','--samples','1',
              '--filesystem-cache','warm,cold-verified','--filesystem-root',str(SCRATCH),
              '--json',str(path)]
        receipt=run(f'qualification-{i:02}',argv)
        rows.append(dict(case=case,receipt=desc(P/f'commands/qualification-{i:02}/receipt.json'),
                         report=desc(path) if path.exists() else None,exit_code=receipt['exit_code']))
    write(P/'qualification.json',dict(rows=rows,status='commands_pass' if all(r['exit_code']==0 for r in rows) else 'failed'))
    assert all(r['exit_code']==0 for r in rows), 'qualification failed; no formal capture'

if __name__=='__main__':
    assert len(sys.argv)==2 and sys.argv[1] in ('prepare','build','quality','qualify')
    globals()[sys.argv[1]]()
