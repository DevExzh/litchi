"""Serial public-constructor malformed CFB guards sharing main source custody."""
from pathlib import Path
import json,shutil
import run as R
G=R.HERE/'guard'

def folder():return G/R.STAGE

def call(name,command,binary=None):
    destination=folder();destination.mkdir(parents=True,exist_ok=True)
    manifest=destination/'source-manifest.json'
    source=R.HERE/R.STAGE/'source-manifest.json'
    if not manifest.exists():shutil.copy2(source,manifest)
    assert R.sha(manifest)==R.sha(source)
    original=R.FOLDER
    try:
        R.FOLDER=destination
        R.run(name,command,binary)
    finally:R.FOLDER=original

def build():
    command=['env','TMPDIR='+str(R.TARGET/'tmp'),'CARGO_BUILD_JOBS=2','CARGO_INCREMENTAL=0','cargo','build','--release','--locked','-p','litchi-cfb','--features','write','--example','perf_chain_guard','--target-dir',str(R.TARGET)]
    call('build-guard',command)
    source=R.TARGET/'release/examples/perf_chain_guard';binary=R.SCRATCH/'guard'
    shutil.copy2(source,binary);assert R.sha(source)==R.sha(binary)
    R.write(folder()/'binary-guard.json',dict(path=str(binary),sha256=R.sha(binary),bytes=binary.stat().st_size,build_receipt_sha256=R.sha(folder()/'build-guard.receipt.json'),source_manifest_sha256=R.sha(R.HERE/R.STAGE/'source-manifest.json')))

def capture(repeat):
    config=json.loads((R.HERE/'plan.json').read_text())
    binary=R.SCRATCH/'guard';assert R.sha(binary)==json.loads((folder()/'binary-guard.json').read_text())['sha256']
    for size in config['guard']['sizes']:
        for case in config['guard']['cases']:
            name=f'guard-r{repeat}-{size}-{case}'
            command=['taskset','-c',str(config['cpu']),str(binary),'--case',case,'--size',str(size),'--warmup',str(config['guard']['warmup']),'--samples',str(config['guard']['samples']),'--json',str(folder()/(name+'.json'))]
            call(name,command,binary)
